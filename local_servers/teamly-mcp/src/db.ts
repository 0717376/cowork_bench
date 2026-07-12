import pg from 'pg';

const { Pool } = pg;

export const pool = new Pool({
  host: process.env.PG_HOST ?? 'localhost',
  port: Number(process.env.PG_PORT ?? 5432),
  database: process.env.PG_DATABASE ?? 'cowork_gym',
  user: process.env.PG_USER ?? 'postgres',
  password: process.env.PG_PASSWORD ?? 'postgres',
});

export async function query<T extends pg.QueryResultRow = any>(text: string, params: any[] = []): Promise<pg.QueryResult<T>> {
  return pool.query<T>(text, params);
}

export type SpaceRow = {
  id: number;
  key: string;
  name: string;
  description: string;
};

export type PageRow = {
  id: number;
  space_id: number;
  space_key: string;
  title: string;
  body: string;
  author: string;
  parent_id: number | null;
  created_at: string;
  updated_at: string;
};

export async function getSpaceByKey(key: string): Promise<SpaceRow | null> {
  const r = await query<SpaceRow>('SELECT * FROM teamly.spaces WHERE key = $1', [key]);
  return r.rows[0] ?? null;
}

export async function findPage(spaceKey: string | undefined, title: string): Promise<PageRow | null> {
  const sql = `
    SELECT p.id, p.space_id, s.key AS space_key, p.title, p.body, p.author,
           p.parent_id, p.created_at::text, p.updated_at::text
      FROM teamly.pages p
      JOIN teamly.spaces s ON s.id = p.space_id
     WHERE LOWER(p.title) = LOWER($1)
       ${spaceKey ? 'AND s.key = $2' : ''}
     LIMIT 1`;
  const params = spaceKey ? [title, spaceKey] : [title];
  const r = await query<PageRow>(sql, params);
  return r.rows[0] ?? null;
}

export async function getLabels(pageId: number): Promise<string[]> {
  const r = await query<{ name: string }>(
    `SELECT l.name FROM teamly.page_labels pl
       JOIN teamly.labels l ON l.id = pl.label_id
      WHERE pl.page_id = $1
      ORDER BY l.name`,
    [pageId],
  );
  return r.rows.map((x) => x.name);
}
