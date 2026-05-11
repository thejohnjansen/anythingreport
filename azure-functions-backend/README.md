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
- `ADO_PAT`
- `GRAPH_ACCESS_TOKEN` (optional fallback; request bearer token is preferred)
- `GRAPH_UPLOAD_MODE` (`delegated` or `application`)
- `GRAPH_UPLOAD_FOLDER`
- `GRAPH_DRIVE_ID` (required for application uploads)
- `GRAPH_USER_ID` (optional delegated upload as specific user)
- `COSMOS_ENDPOINT`
- `COSMOS_KEY`
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

const client = new CosmosClient({
  endpoint: process.env.COSMOS_ENDPOINT!,
  key: process.env.COSMOS_KEY!
});

const container = client
  .database(process.env.COSMOS_DATABASE!)
  .container(process.env.COSMOS_CONTAINER!);
```

## API endpoints

- `GET /api/workitems?iterationPath=...` queries ADO work items under an iteration path.
- `GET /api/workitems?areaPath=...` queries ADO work items under an area path.
- `GET /api/stories?cycle=XXX` returns stored narrative documents.
- `POST /api/stories` creates or updates a story document.
- `PUT /api/stories/{id}` updates a story with a version check.
- `POST /api/stories/export?cycle=XXX` generates a PPT from stories, uploads it to Graph, and returns the file URL.

## Notes

The current implementation uses a simple version field for optimistic concurrency. In a fuller build, you could also persist and compare Cosmos ETags for stricter concurrency control.
```