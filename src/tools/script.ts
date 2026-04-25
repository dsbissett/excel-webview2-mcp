import {zod} from '../third_party/index.js';
import type {Frame, JSHandle, Page, WebWorker} from '../third_party/index.js';
import type {ExtensionServiceWorker} from '../types.js';

import {ToolCategory} from './categories.js';
import type {Context, Response} from './ToolDefinition.js';
import {defineTool, pageIdSchema} from './ToolDefinition.js';
import {ToolError} from './ToolError.js';

export type Evaluatable = Page | Frame | WebWorker;

export const evaluateScript = defineTool(cliArgs => {
  return {
    name: 'evaluate_script',
    description: `Evaluate a JavaScript function inside the currently selected page. Returns the response as JSON,
so returned values have to be JSON-serializable.`,
    annotations: {
      category: ToolCategory.DEBUGGING,
      readOnlyHint: false,
    },
    schema: {
      function: zod.string().describe(
        `A JavaScript function declaration to be executed by the tool in the currently selected page.
Example without arguments: \`() => {
  return document.title
}\` or \`async () => {
  return await fetch("example.com")
}\`.
Example with arguments: \`(el) => {
  return el.innerText;
}\`
`,
      ),
      args: zod
        .array(
          zod
            .string()
            .describe(
              'The uid of an element on the page from the page content snapshot',
            ),
        )
        .optional()
        .describe(`An optional list of arguments to pass to the function.`),
      ...(cliArgs?.experimentalPageIdRouting ? pageIdSchema : {}),
      ...(cliArgs?.categoryExtensions
        ? {
            serviceWorkerId: zod
              .string()
              .optional()
              .describe(
                `An optional service worker id to evaluate the script in.`,
              ),
          }
        : {}),
    },
    handler: async (request, response, context) => {
      const {
        serviceWorkerId,
        args: uidArgs,
        function: fnString,
        pageId,
      } = request.params;

      if (cliArgs?.categoryExtensions && serviceWorkerId) {
        if (uidArgs && uidArgs.length > 0) {
          throw new ToolError({
            category: 'validation',
            isRetryable: false,
            message:
              'args (element uids) cannot be used when evaluating in a service worker.',
            context: {
              toolName: 'evaluate_script',
              attempted: 'evaluate script in service worker',
              failed: 'incompatible parameters',
            },
          });
        }
        if (pageId) {
          throw new ToolError({
            category: 'validation',
            isRetryable: false,
            message: 'specify either a pageId or a serviceWorkerId.',
            context: {
              toolName: 'evaluate_script',
              attempted: 'evaluate script',
              failed: 'mutually exclusive parameters',
            },
          });
        }

        const worker = await getWebWorker(context, serviceWorkerId);
        await context
          .getSelectedMcpPage()
          .waitForEventsAfterAction(async () => {
            await performEvaluation(worker, fnString, [], response);
          });
        return;
      }

      const mcpPage = cliArgs?.experimentalPageIdRouting
        ? context.getPageById(request.params.pageId)
        : context.getSelectedMcpPage();
      const page: Page = mcpPage.pptrPage;

      const args: Array<JSHandle<unknown>> = [];
      try {
        const frames = new Set<Frame>();
        for (const uid of uidArgs ?? []) {
          const handle = await mcpPage.getElementByUid(uid);
          frames.add(handle.frame);
          args.push(handle);
        }

        const evaluatable = await getPageOrFrame(page, frames);

        await mcpPage.waitForEventsAfterAction(async () => {
          await performEvaluation(evaluatable, fnString, args, response);
        });
      } finally {
        void Promise.allSettled(args.map(arg => arg.dispose()));
      }
    },
  };
});

const performEvaluation = async (
  evaluatable: Evaluatable,
  fnString: string,
  args: Array<JSHandle<unknown>>,
  response: Response,
) => {
  const fn = await evaluatable.evaluateHandle(`(${fnString})`);
  try {
    const result = await evaluatable.evaluate(
      async (fn, ...args) => {
        // @ts-expect-error no types for function fn
        return JSON.stringify(await fn(...args));
      },
      fn,
      ...args,
    );
    response.appendResponseLine('Script ran on page and returned:');
    response.appendResponseLine('```json');
    response.appendResponseLine(`${result}`);
    response.appendResponseLine('```');
  } finally {
    void fn.dispose();
  }
};

const getPageOrFrame = async (
  page: Page,
  frames: Set<Frame>,
): Promise<Page | Frame> => {
  let pageOrFrame: Page | Frame;
  // We can't evaluate the element handle across frames
  if (frames.size > 1) {
    throw new Error(
      "Elements from different frames can't be evaluated together.",
    );
  } else {
    pageOrFrame = [...frames.values()][0] ?? page;
  }

  return pageOrFrame;
};

const getWebWorker = async (
  context: Context,
  serviceWorkerId: string,
): Promise<WebWorker> => {
  const serviceWorkers = context.getExtensionServiceWorkers();

  const serviceWorker = serviceWorkers.find(
    (sw: ExtensionServiceWorker) =>
      context.getExtensionServiceWorkerId(sw) === serviceWorkerId,
  );

  if (serviceWorker && serviceWorker.target) {
    const worker = await serviceWorker.target.worker();

    if (!worker) {
      throw new ToolError({
        category: 'not_found',
        isRetryable: false,
        message: 'Service worker target not found.',
        context: {
          toolName: 'evaluate_script',
          attempted: 'resolve service worker target',
          failed: 'service worker target unavailable',
          details: {serviceWorkerId},
        },
      });
    }

    return worker;
  } else {
    throw new ToolError({
      category: 'not_found',
      isRetryable: false,
      message: 'Service worker not found.',
      context: {
        toolName: 'evaluate_script',
        attempted: 'resolve service worker',
        failed: 'service worker id not found',
        details: {serviceWorkerId},
      },
    });
  }
};
