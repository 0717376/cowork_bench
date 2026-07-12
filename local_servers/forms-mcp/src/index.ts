#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
  Request
} from '@modelcontextprotocol/sdk/types.js';
import * as pgForms from './pg-forms.js';

class FormsServer {
  private server: Server;

  constructor() {
    this.server = new Server(
      {
        name: 'forms-mcp',
        version: '0.1.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.setupToolHandlers();

    // Error handling
    this.server.onerror = (error: Error) => console.error('[MCP Error]', error);
    process.on('SIGINT', async () => {
      await this.server.close();
      await pgForms.closePool();
      process.exit(0);
    });
  }

  private setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'list_forms',
          description: 'Получить список всех форм (id, title, description)',
          inputSchema: {
            type: 'object',
            properties: {},
            required: [],
          },
        },
        {
          name: 'create_form',
          description: 'Создать новую форму',
          inputSchema: {
            type: 'object',
            properties: {
              title: {
                type: 'string',
                description: 'Form title',
              },
              description: {
                type: 'string',
                description: 'Form description (optional)',
              }
            },
            required: ['title'],
          },
        },
        {
          name: 'add_text_question',
          description: 'Добавить текстовый вопрос в форму',
          inputSchema: {
            type: 'object',
            properties: {
              formId: {
                type: 'string',
                description: 'Form ID',
              },
              questionTitle: {
                type: 'string',
                description: 'Question title',
              },
              required: {
                type: 'boolean',
                description: 'Whether required (optional, default is false)',
              }
            },
            required: ['formId', 'questionTitle'],
          },
        },
        {
          name: 'add_multiple_choice_question',
          description: 'Добавить вопрос с вариантами ответа в форму',
          inputSchema: {
            type: 'object',
            properties: {
              formId: {
                type: 'string',
                description: 'Form ID',
              },
              questionTitle: {
                type: 'string',
                description: 'Question title',
              },
              options: {
                type: 'array',
                items: {
                  type: 'string'
                },
                description: 'Array of choices',
              },
              required: {
                type: 'boolean',
                description: 'Whether required (optional, default is false)',
              }
            },
            required: ['formId', 'questionTitle', 'options'],
          },
        },
        {
          name: 'get_form',
          description: 'Получить форму со списком вопросов',
          inputSchema: {
            type: 'object',
            properties: {
              formId: {
                type: 'string',
                description: 'Form ID',
              }
            },
            required: ['formId'],
          },
        },
        {
          name: 'get_form_responses',
          description: 'Получить ответы на форму',
          inputSchema: {
            type: 'object',
            properties: {
              formId: {
                type: 'string',
                description: 'Form ID',
              }
            },
            required: ['formId'],
          },
        }
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request: any) => {
      try {
        switch (request.params.name) {
          case 'list_forms':
            return await this.listForms();
          case 'create_form':
            return await this.createForm(request.params.arguments);
          case 'add_text_question':
            return await this.addTextQuestion(request.params.arguments);
          case 'add_multiple_choice_question':
            return await this.addMultipleChoiceQuestion(request.params.arguments);
          case 'get_form':
            return await this.getForm(request.params.arguments);
          case 'get_form_responses':
            return await this.getFormResponses(request.params.arguments);
          default:
            throw new McpError(
              ErrorCode.MethodNotFound,
              `Unknown tool: ${request.params.name}`
            );
        }
      } catch (error: any) {
        console.error('Error in tool execution:', error);
        return {
          content: [
            {
              type: 'text',
              text: `Error: ${error.message || 'Unknown error'}`,
            },
          ],
          isError: true,
        };
      }
    });
  }

  private async listForms() {
    try {
      const data = await pgForms.listForms();
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(data, null, 2),
          },
        ],
      };
    } catch (error: any) {
      console.error('Error listing forms:', error);
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to list forms: ${error.message}`
      );
    }
  }

  private async createForm(args: any) {
    if (!args.title) {
      throw new McpError(ErrorCode.InvalidParams, 'Title is required');
    }

    try {
      const result = await pgForms.createForm(args.title, args.description);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error: any) {
      console.error('Error creating form:', error);
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to create form: ${error.message}`
      );
    }
  }

  private async addTextQuestion(args: any) {
    if (!args.formId || !args.questionTitle) {
      throw new McpError(
        ErrorCode.InvalidParams,
        'Form ID and question title are required'
      );
    }

    try {
      const result = await pgForms.addTextQuestion(args.formId, args.questionTitle, args.required || false);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error: any) {
      console.error('Error adding text question:', error);
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to add text question: ${error.message}`
      );
    }
  }

  private async addMultipleChoiceQuestion(args: any) {
    if (!args.formId || !args.questionTitle || !args.options || !Array.isArray(args.options)) {
      throw new McpError(
        ErrorCode.InvalidParams,
        'Form ID, question title, and options array are required'
      );
    }

    try {
      const result = await pgForms.addMultipleChoiceQuestion(
        args.formId, args.questionTitle, args.options, args.required || false
      );
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error: any) {
      console.error('Error adding multiple choice question:', error);
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to add multiple choice question: ${error.message}`
      );
    }
  }

  private async getForm(args: any) {
    if (!args.formId) {
      throw new McpError(ErrorCode.InvalidParams, 'Form ID is required');
    }

    try {
      const formData = await pgForms.getForm(args.formId);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(formData, null, 2),
          },
        ],
      };
    } catch (error: any) {
      console.error('Error getting form:', error);
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to get form: ${error.message}`
      );
    }
  }

  private async getFormResponses(args: any) {
    if (!args.formId) {
      throw new McpError(ErrorCode.InvalidParams, 'Form ID is required');
    }

    try {
      const responseData = await pgForms.getFormResponses(args.formId);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(responseData, null, 2),
          },
        ],
      };
    } catch (error: any) {
      console.error('Error getting form responses:', error);
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to get form responses: ${error.message}`
      );
    }
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Forms MCP server running on stdio');
  }
}

const server = new FormsServer();
server.run().catch(console.error);
