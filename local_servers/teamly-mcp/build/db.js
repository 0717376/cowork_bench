import pg from 'pg';
const { Pool } = pg;
export const pool = new Pool({
    host: process.env.PG_HOST ?? 'localhost',
    port: Number(process.env.PG_PORT ?? 5432),
    database: process.env.PG_DATABASE ?? 'cowork_gym',
    user: process.env.PG_USER ?? 'postgres',
    password: process.env.PG_PASSWORD ?? 'postgres',
});
export async function query(text, params = []) {
    return pool.query(text, params);
}
export async function getSpaceByKey(key) {
    const r = await query('SELECT * FROM teamly.spaces WHERE key = $1', [key]);
    return r.rows[0] ?? null;
}
export async function findPage(spaceKey, title) {
    const sql = `
    SELECT p.id, p.space_id, s.key AS space_key, p.title, p.body, p.author,
           p.parent_id, p.created_at::text, p.updated_at::text
      FROM teamly.pages p
      JOIN teamly.spaces s ON s.id = p.space_id
     WHERE LOWER(p.title) = LOWER($1)
       ${spaceKey ? 'AND s.key = $2' : ''}
     LIMIT 1`;
    const params = spaceKey ? [title, spaceKey] : [title];
    const r = await query(sql, params);
    return r.rows[0] ?? null;
}
export async function getLabels(pageId) {
    const r = await query(`SELECT l.name FROM teamly.page_labels pl
       JOIN teamly.labels l ON l.id = pl.label_id
      WHERE pl.page_id = $1
      ORDER BY l.name`, [pageId]);
    return r.rows.map((x) => x.name);
}
