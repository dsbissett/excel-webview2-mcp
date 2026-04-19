import {zod, PredefinedNetworkConditions} from '../third_party/index.js';

import {ToolCategory} from './categories.js';
import {definePageTool} from './ToolDefinition.js';

const throttlingOptions: [string, ...string[]] = [
  'Offline',
  ...Object.keys(PredefinedNetworkConditions),
];

export const emulate = definePageTool({
  name: 'emulate',
  description: `Throttles network and/or CPU on the selected page.`,
  annotations: {
    category: ToolCategory.EMULATION,
    readOnlyHint: false,
  },
  schema: {
    networkConditions: zod
      .enum(throttlingOptions)
      .optional()
      .describe(`Throttle network. Omit to disable throttling.`),
    cpuThrottlingRate: zod
      .number()
      .min(1)
      .max(20)
      .optional()
      .describe(
        'Represents the CPU slowdown factor. Omit or set the rate to 1 to disable throttling',
      ),
  },
  handler: async (request, _response, context) => {
    const page = request.page;
    await context.emulate(request.params, page.pptrPage);
  },
});
