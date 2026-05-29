/**
 * Vercel serverless — menyajikan konfigurasi client dari environment variables.
 */
module.exports = (req, res) => {
  const config = {
    SUPABASE_URL: process.env.SUPABASE_URL || '',
    SUPABASE_ANON_KEY: process.env.SUPABASE_ANON_KEY || '',
  };
  res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.status(200).send(`window.APP_CONFIG = ${JSON.stringify(config)};`);
};
