import { createSendTeamsMessageViaWebhookAction } from './send-ms-teams-message';
import { createMockActionContext } from '@backstage/plugin-scaffolder-node-test-utils';
import axios from 'axios';
import { Config } from '@backstage/config';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios>;

describe('ms-teams:sendMessage', () => {
  afterEach(() => {
    jest.resetAllMocks();
  });

  it('should throw error if webhookUrl is not defined', async () => {
    const action = createSendTeamsMessageViaWebhookAction({
      config: {
        getOptionalString: (_key: string) => undefined,
      } as Config,
    });

    await expect(
      action.handler(
        createMockActionContext({
          input: {
            message: 'Hello, Teams!',
          },
        }),
      ),
    ).rejects.toThrow(
      'Webhook URL is not specified in either the action input or the app-config. This must be specified in at least one place in order to send a message',
    );
  });

  it('should send to config webhook URL if provided', async () => {
    const action = createSendTeamsMessageViaWebhookAction({
      config: {
        getOptionalString: (_key: string) => 'https://example-teams.com',
      } as Config,
    });

    mockedAxios.post.mockResolvedValue({ status: 200 });

    await action.handler(
      createMockActionContext({
        input: {
          message: 'Hello, Teams!',
        },
      }),
    );

    expect(mockedAxios.post).toHaveBeenCalledWith(
      'https://example-teams.com',
      expect.anything(),
    );
  });

  it('should prefer webhook url from input over config if provided', async () => {
    const action = createSendTeamsMessageViaWebhookAction({
      config: {
        getOptionalString: (_key: string) => 'https://config-teams.com',
      } as Config,
    });

    mockedAxios.post.mockResolvedValue({ status: 200 });

    await action.handler(
      createMockActionContext({
        input: {
          message: 'Hello, Teams!',
          webhookUrl: 'https://input-teams.com',
        },
      }),
    );

    // webhookUrl from input takes precedence now
    expect(mockedAxios.post).toHaveBeenCalledWith(
      'https://input-teams.com',
      expect.anything(),
    );
  });

  it('should use webhookUrl from input if config value is not present', async () => {
    const action = createSendTeamsMessageViaWebhookAction({
      config: {
        getOptionalString: (_key: string) => undefined,
      } as Config,
    });

    mockedAxios.post.mockResolvedValue({ status: 200 });

    await action.handler(
      createMockActionContext({
        input: {
          message: 'Hello, Teams!',
          webhookUrl: 'https://nevergonnagiveyouup-teams.com',
        },
      }),
    );

    expect(mockedAxios.post).toHaveBeenCalledWith(
      'https://nevergonnagiveyouup-teams.com',
      expect.anything(),
    );
  });

  it('should send message in proper format to webhook URL', async () => {
    const action = createSendTeamsMessageViaWebhookAction({
      config: {
        getOptionalString: (_key: string) => 'https://example-teams.com',
      } as Config,
    });

    mockedAxios.post.mockResolvedValue({ status: 200 });

    await action.handler(
      createMockActionContext({
        input: {
          message: 'Hello, Teams!',
          webhookUrl: 'https://specific-url.com',
        },
      }),
    );

    expect(mockedAxios.post).toHaveBeenCalledWith(
      'https://specific-url.com',
      expect.objectContaining({
        attachments: [
          {
            contentType: 'application/vnd.microsoft.card.adaptive',
            contentUrl: null,
            content: expect.objectContaining({ body: [{ type: 'TextBlock', text: 'Hello, Teams!', wrap: true }] }),
          },
        ],
      }),
    );
  });

  it('should throw an error if result.status is not 200', async () => {
    const action = createSendTeamsMessageViaWebhookAction({
      config: {
        getOptionalString: (_key: string) => 'https://example-teams.com',
      } as Config,
    });

    mockedAxios.post.mockResolvedValue({ status: 400 });

    await expect(
      action.handler(
        createMockActionContext({
          input: {
            message: 'Hello, Teams!',
            webhookUrl: 'https://input-failure.com',
          },
        }),
      ),
    ).rejects.toThrow(
      'Something went wrong while trying to send a request to the Teams webhook URL - StatusCode 400',
    );
  });

  it('should throw an error if result.status is not 202', async () => {
    const action = createSendTeamsMessageViaWebhookAction({
      config: {
        getOptionalString: (_key: string) => 'https://example-teams.com',
      } as Config,
    });

    mockedAxios.post.mockResolvedValue({ status: 400 });

    await expect(
      action.handler(
        createMockActionContext({
          input: {
            message: 'Hello, Teams!',
            webhookUrl: 'https://input-failure.com',
          },
        }),
      ),
    ).rejects.toThrow(
      'Something went wrong while trying to send a request to the Teams webhook URL - StatusCode 400',
    );
  });
});
