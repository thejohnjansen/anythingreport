# Azure Functions backend scaffold

Minimal Azure Functions backend for multi-user storytelling around Azure DevOps work items.

## Layout

```
azure-functions-backend/
  host.json
  local.settings.json
  package.json
  tsconfig.json
  src/
    index.ts
    types.ts
    cosmosClient.ts
    storyRepository.ts
    workItemService.ts
    api/
      stories.ts
      workitems.ts
```

## Environment variables

- `ADO_ORGANIZATION_URL`
- `ADO_PROJECT`
- `ADO_PAT` (optional fallback when Azure CLI auth is unavailable)
- `GRAPH_ACCESS_TOKEN` (optional fallback; request bearer token is preferred)
- `GRAPH_UPLOAD_MODE` (`delegated` or `application`)
- `GRAPH_UPLOAD_FOLDER`
- `GRAPH_DRIVE_ID` (required for application uploads)
- `GRAPH_USER_ID` (optional delegated upload as specific user)
- `COSMOS_CONNECTION_STRING`
- `COSMOS_DATABASE`
- `COSMOS_CONTAINER`

## Data model

```ts
{
  workItemId: number;
  cycle: string;
  title: string;
  area: string;
  pmComments: string;
  narrative: string;
  lastEditedBy: string;
  lastEditedAt: string;
  version: number;
}
```

## Cosmos example

```ts
import { CosmosClient } from '@azure/cosmos';

const client = new CosmosClient(process.env.COSMOS_CONNECTION_STRING!);

const container = client
  .database(process.env.COSMOS_DATABASE!)
  .container(process.env.COSMOS_CONTAINER!);
```

## API endpoints

- `POST /api/slides` mirrors the existing root app preview flow: accepts `queryUrl`, runs the ADO tree query, and returns grouped slide data for the frontend preview.
- `POST /api/slides/export` accepts preview state plus top-of-mind and pipeline data and returns a `.pptx` buffer using the existing report layout logic.
- `POST /api/doc` accepts `queryUrl` and returns the document sections used by the doc preview.
- `POST /api/retro` accepts `queryUrl` and returns the retro rows used by the retro preview.
- `GET /api/workitems?iterationPath=...` queries ADO work items under an iteration path.
- `GET /api/workitems?areaPath=...` queries ADO work items under an area path.
- `GET /api/stories?cycle=XXX` returns stored narrative documents.
- `POST /api/stories` creates or updates a story document.
- `PUT /api/stories/{id}` updates a story with a version check.
- `POST /api/stories/export?cycle=XXX` generates a PPT from stories, uploads it to Graph, and returns the file URL.

## Local seed data

After pointing `local.settings.json` at a real Cosmos account or the local Cosmos emulator, you can seed five sample story documents with:

```bash
npm run seed:stories
```

The seed script creates the configured database and `stories` container if they do not exist, using `/cycle` as the partition key.

## Notes

The current implementation uses a simple version field for optimistic concurrency. In a fuller build, you could also persist and compare Cosmos ETags for stricter concurrency control.

For local ADO access, the backend now prefers `az account get-access-token` and only falls back to `ADO_PAT` if Azure CLI auth is unavailable.
```