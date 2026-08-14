// Shared output contract: every office tool returns a single text block.

import type { ContentBlock } from '@deepseek-ai/dsh-llm'

export interface TextOutput {
  schema: {
    type: 'object'
    additionalProperties: boolean
    properties: {
      content: { type: 'string'; required: true; description: string }
    }
  }
  render: (args: unknown, value: { content: string }) => ContentBlock[]
}

export const textOutput: TextOutput = {
  schema: {
    type: 'object',
    additionalProperties: false,
    properties: {
      content: { type: 'string', required: true, description: 'Human-readable tool result.' },
    },
  },
  render: (_args, value) => [{ type: 'text', text: value.content }],
}
