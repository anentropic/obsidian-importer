import { BatchInterceptor } from '@mswjs/interceptors';
import { FetchInterceptor } from '@mswjs/interceptors/fetch';

/**
 * Request handler type for mocking HTTP requests
 */
export type RequestHandler = (url: string | URL) => Promise<Response> | Response;

/**
 * Creates a request interceptor using MSW interceptors.
 * This is the recommended approach for mocking requests in Vitest for Node.js.
 */
export class RequestMock {
	private interceptor: BatchInterceptor;
	private handlers: RequestHandler[] = [];

	constructor() {
		// Create a batch interceptor with fetch support
		this.interceptor = new BatchInterceptor({
			name: 'test-interceptor',
			interceptors: [new FetchInterceptor()],
		});
	}

	/**
	 * Register a request handler
	 */
	addHandler(handler: RequestHandler): void {
		this.handlers.push(handler);
	}

	/**
	 * Start intercepting requests
	 */
	start(): void {
		this.interceptor.apply();

		this.interceptor.on('request', async ({ request, requestId, controller }) => {
			const url = new URL(request.url);

			// Try each handler in order
			for (const handler of this.handlers) {
				try {
					const response = await handler(url);
					if (response) {
						// Create a proper fetch Response from the handler response
						const responseInit: ResponseInit = {
							status: response.status,
							statusText: response.statusText,
							headers: response.headers,
						};
						
						// Clone and read the body to avoid consuming the original stream
						const body = await response.clone().arrayBuffer();
						const mockResponse = new Response(body, responseInit);
						
						// Respond with the mocked response using controller
						controller.respondWith(mockResponse);
						return;
					}
				}
				catch (error) {
					// Handler threw an error or returned a non-matching response,
					// continue to next handler to allow fallback behavior
					continue;
				}
			}

			// No handler matched, return 404
			controller.respondWith(new Response('Not Found', { status: 404 }));
		});
	}

	/**
	 * Stop intercepting requests and clean up
	 */
	stop(): void {
		this.interceptor.dispose();
	}

	/**
	 * Clear all registered handlers
	 */
	clearHandlers(): void {
		this.handlers = [];
	}

	/**
	 * Reset the interceptor (stop, clear handlers, and restart)
	 */
	reset(): void {
		this.stop();
		this.clearHandlers();
		// Create a new interceptor
		this.interceptor = new BatchInterceptor({
			name: 'test-interceptor',
			interceptors: [new FetchInterceptor()],
		});
	}
}

/**
 * Creates a simple request mock for testing
 */
export function createRequestMock(): RequestMock {
	return new RequestMock();
}
