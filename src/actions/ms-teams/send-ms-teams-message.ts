import { createTemplateAction } from '@backstage/plugin-scaffolder-node';
import axios from 'axios';
import { Config } from '@backstage/config';
import { InputError } from '@backstage/errors';
import { examples } from './send-ms-teams-message.examples';

/**
 * Creates a `ms-teams:sendMessage` Scaffolder action.
 *
 * @public
 */
export function createSendTeamsMessageViaWebhookAction(options: { config: Config }) {
  const { config } = options;

  return createTemplateAction({
    id: 'ms-teams:sendMessage',
    description: 'Sends a Microsoft Teams message via a webhook',
    examples,
    schema: {
      input: z =>
        z.object({
          message: z.string().describe('The message to send via webhook'),
          webhookUrl: z
            .string()
            .describe(
              'The Microsoft Teams webhook URL to send the request to. The URL must either be specified here or in the Backstage config',
            )
            .optional(),
        }),
    },
    async handler(ctx) {
      const webhookUrl =
        ctx.input.webhookUrl ?? config.getOptionalString('ms-teams.webhookUrl');

      if (!webhookUrl) {
        throw new InputError(
          'Webhook URL is not specified in either the action input or the app-config. This must be specified in at least one place in order to send a message',
        );
      }

      const body = {
        "@type": "MessageCard",
        "@context": "http://schema.org",
        "summary": "Backstage",
        "text": ctx.input.message ?? "",
      };

      const result = await axios.post(webhookUrl, body);
      const isSuccess = result.status === 200 || result.status === 202;
      if (!isSuccess) {
        ctx.logger.error(
          `Something went wrong while trying to send a request to the Teams webhook URL - StatusCode ${result.status}`,
        );
        ctx.logger.debug(`Response body: ${result.data}`);
        ctx.logger.debug(`Webhook URL: ${webhookUrl}`);
        ctx.logger.debug(`Input message: ${ctx.input.message}`);
        throw new Error(
          `Something went wrong while trying to send a request to the Teams webhook URL - StatusCode ${result.status}`,
        );
      }
    },
  });
}
