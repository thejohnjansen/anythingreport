import type { Container } from '@azure/cosmos';
import { getCosmosContainer } from './cosmosClient';
import type { BoardDocument, BoardUpsertRequest } from './types';

const BOARD_PARTITION = '__board__';

export function buildBoardId(boardType: string, teamSlug: string): string {
  return `${boardType}:${teamSlug}`;
}

export class BoardRepository {
  constructor(private readonly container: Container = getCosmosContainer()) {}

  async getBoardDocument(boardType: string, teamSlug: string): Promise<BoardDocument | null> {
    const id = buildBoardId(boardType, teamSlug);
    const query = {
      query: 'SELECT TOP 1 * FROM c WHERE c.id = @id AND c.cycle = @cycle',
      parameters: [
        { name: '@id', value: id },
        { name: '@cycle', value: BOARD_PARTITION }
      ]
    };
    const { resources } = await this.container.items.query<BoardDocument>(query).fetchAll();
    return resources[0] ?? null;
  }

  async upsertBoardDocument(req: BoardUpsertRequest): Promise<BoardDocument> {
    const id = buildBoardId(req.boardType, req.teamSlug);
    const existing = await this.getBoardDocument(req.boardType, req.teamSlug);

    if (typeof req.expectedVersion === 'number' && existing && req.expectedVersion !== existing.version) {
      const err = Object.assign(
        new Error(`Version conflict for board ${id}. Expected ${req.expectedVersion}, found ${existing.version}.`),
        { code: 'CONFLICT', currentDocument: existing }
      );
      throw err;
    }

    const now = new Date().toISOString();
    const document: BoardDocument = {
      id,
      boardType: req.boardType,
      teamSlug: req.teamSlug,
      cycle: BOARD_PARTITION,
      content: req.content,
      lastEditedBy: req.lastEditedBy,
      lastEditedAt: now,
      version: (existing?.version ?? 0) + 1
    };

    const response = await this.container.items.upsert<BoardDocument>(document);
    if (!response.resource) {
      throw new Error('Cosmos upsert returned no resource.');
    }
    return response.resource;
  }
}
