const { Pool } = require('pg');

async function run() {
  const pool = new Pool({ connectionString: process.env.DATABASE_URL });
  const client = await pool.connect();
  try {
    const res = await client.query('SELECT count(*) FROM facilities');
    console.log(`✅ Keep-alive ping OK — facilities table has ${res.rows[0].count} rows.`);
  } finally {
    client.release();
    await pool.end();
  }
}

run().catch((err) => {
  console.error('❌ Keep-alive ping failed:', err.message);
  process.exit(1);
});
