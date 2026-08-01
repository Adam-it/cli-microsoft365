import * as http from 'http';
import { IncomingMessage, ServerResponse } from 'http';
import { AddressInfo } from 'net';
import { ParsedUrlQuery } from 'querystring';
import * as fs from 'fs';
import * as path from 'path';
import * as url from "url";
import { Auth, InteractiveAuthorizationCodeResponse, Service, InteractiveAuthorizationErrorResponse } from './Auth';
import { Logger } from './cli/Logger';
import { browserUtil } from './utils/browserUtil';

export class AuthServer {
  // assigned through this.initializeServer() hence !
  private httpServer!: http.Server;
  // assigned through this.initializeServer() hence !
  private service!: Service;
  // assigned through this.initializeServer() hence !
  private resolve!: (error: InteractiveAuthorizationCodeResponse) => void;
  // assigned through this.initializeServer() hence !
  private reject!: (error: InteractiveAuthorizationErrorResponse) => void;
  // assigned through this.initializeServer() hence !
  private logger!: Logger;

  private debug: boolean = false;
  private resource: string = "";
  private generatedServerUrl: string = "";

  public get server(): http.Server {
    return this.httpServer;
  }

  public initializeServer = (service: Service, resource: string, resolve: (result: InteractiveAuthorizationCodeResponse) => void, reject: (error: InteractiveAuthorizationErrorResponse) => void, logger: Logger, debug: boolean = false): void => {
    this.service = service;
    this.resolve = resolve;
    this.reject = reject;
    this.logger = logger;
    this.debug = debug;
    this.resource = resource;

    this.httpServer = http.createServer(this.httpRequest).listen(0, this.httpListener);
  };

  private getAssetDataUri(assetFileName: string): string {
    try {
      const imagePath = path.join(__dirname, 'assets', assetFileName);
      const imageBuffer = fs.readFileSync(imagePath);
      const mimeType = path.extname(assetFileName).toLowerCase() === '.png' ? 'image/png' : 'application/octet-stream';
      return `data:${mimeType};base64,${imageBuffer.toString('base64')}`;
    }
    catch {
      return '';
    }
  }

  private httpListener = (): void => {
    const requestState = Math.random().toString(16).substr(2, 20);
    const address = this.httpServer.address() as AddressInfo;
    this.generatedServerUrl = `http://localhost:${address.port}`;
    const url = `${Auth.getEndpointForResource('https://login.microsoftonline.com', this.service.cloudType)}/${this.service.tenant}/oauth2/authorize?response_type=code&client_id=${this.service.appId}&redirect_uri=${this.generatedServerUrl}&state=${requestState}&resource=${this.resource}&prompt=select_account`;
    if (this.debug) {
      this.logger.logToStderr('Redirect URL:');
      this.logger.logToStderr(url);
      this.logger.logToStderr('');
    }
    this.openUrl(url);
  };

  private openUrl(url: string): void {
    browserUtil.open(url)
      .then(_ => {
        this.logger.log("To sign in, use the web browser that just has been opened. Please sign-in there.");
      })
      .catch(_ => {
        const errorResponse: InteractiveAuthorizationErrorResponse = {
          error: "Can't open the default browser",
          errorDescription: "Was not able to open a browser instance. Try again later or use a different authentication method."
        };

        this.reject(errorResponse);
        this.httpServer.close();
      });
  }

  private httpRequest = (request: IncomingMessage, response: ServerResponse): void => {
    if (this.debug) {
      this.logger.logToStderr('Response:');
      this.logger.logToStderr(request.url);
      this.logger.logToStderr('');
    }

    // url.parse is deprecated but we can't move to URL, because it doesn't
    // support server-relative URLs
    const queryString: ParsedUrlQuery = url.parse(request.url as string, true).query;
    const hasCode: boolean = queryString.code !== undefined;
    const hasError: boolean = queryString.error !== undefined;

    let body: string = "";
    if (hasCode === true) {
      const toolkitLogoImage = this.getAssetDataUri('logo-large.png');
      const pnpLogoImage = this.getAssetDataUri('pnp-logo.png');
      const toolkitLogoImageTag = toolkitLogoImage !== '' ? `<img class="logo-large" src="${toolkitLogoImage}" alt="SPFx Toolkit logo" />` : '';
      const pnpLogoImageTag = pnpLogoImage !== '' ? `<img class="pnp-logo" src="${pnpLogoImage}" alt="PnP logo" />` : '';

      body = `
        <script type="text/JavaScript">
          setTimeout(function(){ window.location = "https://spfxtoolkit.community.ms/"; }, 10000);
        </script>
        <div class="page">
          <div class="card">
            ${toolkitLogoImageTag}
            <h1>Signed in to SPFx Toolkit</h1>
            <p>You can close this window and return to VS Code.</p>
            <p>This page will redirect you to the <a href="https://spfxtoolkit.community.ms/">SPFx Toolkit</a> documentation in a few seconds.</p>
            <div class="community">
              ${pnpLogoImageTag}
              <p>This extension is supported by Microsoft 365 &amp; Power Platform Community.</p>
              <p><a href="https://pnp.github.io/">Microsoft 365 &amp; Power Platform Community</a></p>
            </div>
          </div>
        </div>`;

      this.resolve(<InteractiveAuthorizationCodeResponse>{
        code: queryString.code as string,
        redirectUri: this.generatedServerUrl
      });
    }

    if (hasError === true) {
      const errorMessage: InteractiveAuthorizationErrorResponse = {
        error: queryString.error as string,
        errorDescription: queryString.error_description as string
      };

      body = "<p>Oops! Entra ID replied with an error message.</p>";
      body += `<p>${errorMessage.error}</p>`;
      if (errorMessage.errorDescription !== undefined) {
        body += `<p>${errorMessage.errorDescription}</p>`;
      }

      this.reject(errorMessage);
    }

    if (hasCode === false && hasError === false) {
      const errorMessage: InteractiveAuthorizationErrorResponse = {
        error: "invalid request",
        errorDescription: "An invalid request has been received by the HTTP server"
      };

      body = "<p>Oops! This is an invalid request.</p>";
      body += `<p>${errorMessage.error}</p>`;
      body += `<p>${errorMessage.errorDescription}</p>`;

      this.reject(errorMessage);
    }

    response.writeHead(200, { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'text/html' });
    response.write(`<html><head><meta charset="utf-8" /><meta name="viewport" content="width=device-width, initial-scale=1" /><title>SPFx Toolkit</title><style>
      :root {
        color-scheme: dark;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      body {
        margin: 0;
        min-height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
        background: #1f1f1f;
        color: #cccccc;
      }
      .page {
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 24px;
      }
      .card {
        width: min(480px, 100%);
        background: #252526;
        border: 1px solid #3c3c3c;
        border-radius: 8px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.24);
        padding: 32px;
        text-align: center;
        display: flex;
        flex-direction: column;
      }
      .logo-large {
        width: min(260px, 100%);
        height: auto;
        margin: 0 auto 20px;
      }
      h1 {
        margin: 0 0 8px;
        color: #ffffff;
        font-size: 24px;
        font-weight: 600;
      }
      p {
        margin: 0 0 12px;
        line-height: 1.5;
      }
      a {
        color: #3794ff;
        text-decoration: none;
      }
      a:hover {
        text-decoration: underline;
      }
      .community {
        margin-top: 20px;
        padding-top: 16px;
        border-top: 1px solid #3c3c3c;
      }
      .community p {
        margin: 0 0 8px;
      }
      .pnp-logo {
        width: 56px;
        height: auto;
        margin: 0 auto 8px;
        display: block;
      }
    </style></head><body>${body}</body></html>`);
    response.end();

    this.httpServer.close();
  };
}

export default new AuthServer();