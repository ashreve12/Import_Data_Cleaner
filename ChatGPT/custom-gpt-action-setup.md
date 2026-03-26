# Custom GPT Action Setup

## 1. Deploy the backend

ChatGPT Actions cannot call `http://127.0.0.1:8000` or any other localhost URL.
Deploy the FastAPI app to a public HTTPS URL first, then replace the placeholder
server URL in `action-openapi.yaml`. The server origin must exactly match your deployed HTTPS origin, for example `https://wafer-cleaner-api-production.up.railway.app`.

This deployment can now serve both:

- the Custom GPT Action route at `/run-wafer-cleaner`
- the browser upload page at `/wafer-cleaner`

## 2. Use no authentication first

In the GPT Action builder, select `None` for authentication unless you add an API
key or OAuth layer to your deployed service.

## 3. Paste the schema

Open the GPT builder, add an Action, and paste the contents of `action-openapi.yaml`.
Update:

- `https://wafer-cleaner-api-production.up.railway.app`

with your real deployed base URL.

## 4. Recommended GPT instructions

Tell the GPT to:

- ask the user to upload exactly one Excel workbook
- ask for an optional zip filename if the user wants a custom bundle name
- call `runWaferCleaner` after the workbook is uploaded
- return both the cleaned workbook and the zip bundle to the user
- explain that the user's browser or device controls the final local download location

## 5. Test both front ends

Before using the Action in ChatGPT, test the deployed endpoint in Postman or another
API client. You can also open the browser page at `/wafer-cleaner` to confirm the
HTML upload flow works against the same deployment.
