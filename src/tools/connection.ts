/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import type {ParsedArguments} from '../bin/excel-webview2-mcp-cli-options.js';
import {
  getConnectionStatusSnapshot,
  refreshConnectionStatusProbe,
} from '../connection/status.js';
import {zod} from '../third_party/index.js';

import {ToolCategory} from './categories.js';
import {defineTool} from './ToolDefinition.js';

export const connectionStatus = defineTool((args?: ParsedArguments) => ({
  name: 'connection_status',
  description:
    'Reports whether the server is currently attached to a browser and which CDP endpoint it is tracking.',
  annotations: {
    category: ToolCategory.DEBUGGING,
    readOnlyHint: true,
  },
  requiresContext: false,
  schema: {
    probe: zod
      .boolean()
      .optional()
      .describe(
        'If true, re-runs the CDP /json/version probe for the tracked endpoint instead of returning cached probe state.',
      ),
  },
  handler: async (request, response) => {
    const status = request.params.probe
      ? await refreshConnectionStatusProbe(args?.connectTimeout ?? 5000)
      : getConnectionStatusSnapshot();

    response.setStructuredContent(status);
    response.appendResponseLine(JSON.stringify(status, null, 2));
  },
}));
