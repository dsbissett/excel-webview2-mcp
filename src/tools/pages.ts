/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import {logger} from '../logger.js';
import type {CdpPage} from '../third_party/index.js';
import {zod} from '../third_party/index.js';

import {ToolCategory} from './categories.js';
import {
  CLOSE_PAGE_ERROR,
  definePageTool,
  defineTool,
} from './ToolDefinition.js';

export const listPages = defineTool(args => {
  return {
    name: 'list_pages',
    description: `Get a list of pages ${args?.categoryExtensions ? 'including extension service workers' : ''} open in the browser.`,
    annotations: {
      category: ToolCategory.NAVIGATION,
      readOnlyHint: true,
    },
    schema: {},
    handler: async (_request, response) => {
      response.setIncludePages(true);
      response.setListInPageTools();
    },
  };
});

export const selectPage = defineTool({
  name: 'select_page',
  description: `Select a page as a context for future tool calls.`,
  annotations: {
    category: ToolCategory.NAVIGATION,
    readOnlyHint: true,
  },
  schema: {
    pageId: zod
      .number()
      .describe(
        `The ID of the page to select. Call ${listPages().name} to get available pages.`,
      ),
    bringToFront: zod
      .boolean()
      .optional()
      .describe('Whether to focus the page and bring it to the top.'),
  },
  handler: async (request, response, context) => {
    const page = context.getPageById(request.params.pageId);
    context.selectPage(page);
    response.setIncludePages(true);
    response.setListInPageTools();
    if (request.params.bringToFront) {
      await page.pptrPage.bringToFront();
    }
  },
});

export const closePage = defineTool({
  name: 'close_page',
  description: `Closes the page by its index. The last open page cannot be closed.`,
  annotations: {
    category: ToolCategory.NAVIGATION,
    readOnlyHint: false,
  },
  schema: {
    pageId: zod
      .number()
      .describe('The ID of the page to close. Call list_pages to list pages.'),
  },
  handler: async (request, response, context) => {
    try {
      await context.closePage(request.params.pageId);
    } catch (err) {
      if (err.message === CLOSE_PAGE_ERROR) {
        response.appendResponseLine(err.message);
      } else {
        throw err;
      }
    }
    response.setIncludePages(true);
    response.setListInPageTools();
  },
});

export const handleDialog = definePageTool({
  name: 'handle_dialog',
  description: `If a browser dialog was opened, use this command to handle it`,
  annotations: {
    category: ToolCategory.INPUT,
    readOnlyHint: false,
  },
  schema: {
    action: zod
      .enum(['accept', 'dismiss'])
      .describe('Whether to dismiss or accept the dialog'),
    promptText: zod
      .string()
      .optional()
      .describe('Optional prompt text to enter into the dialog.'),
  },
  handler: async (request, response, _context) => {
    const page = request.page;
    const dialog = page.getDialog();
    if (!dialog) {
      throw new Error('No open dialog found');
    }

    switch (request.params.action) {
      case 'accept': {
        try {
          await dialog.accept(request.params.promptText);
        } catch (err) {
          logger(err);
        }
        response.appendResponseLine('Successfully accepted the dialog');
        break;
      }
      case 'dismiss': {
        try {
          await dialog.dismiss();
        } catch (err) {
          logger(err);
        }
        response.appendResponseLine('Successfully dismissed the dialog');
        break;
      }
    }

    page.clearDialog();
    response.setIncludePages(true);
  },
});

export const getTabId = definePageTool({
  name: 'get_tab_id',
  description: `Get the tab ID of the page`,
  annotations: {
    category: ToolCategory.NAVIGATION,
    readOnlyHint: true,
    conditions: ['experimentalInteropTools'],
  },
  schema: {
    pageId: zod
      .number()
      .describe(
        `The ID of the page to get the tab ID for. Call ${listPages().name} to get available pages.`,
      ),
  },
  handler: async (request, response, context) => {
    const page = context.getPageById(request.params.pageId);
    const tabId = (page.pptrPage as unknown as CdpPage)._tabId;
    response.setTabId(tabId);
  },
});
