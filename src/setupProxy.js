const { createProxyMiddleware } = require("http-proxy-middleware");

module.exports = function (app) {
  app.use(
    "/api/claude",
    createProxyMiddleware({
      target: "https://api.anthropic.com",
      changeOrigin: true,
      secure: false,
      pathRewrite: { "^/api/claude": "" },
      on: {
        error: (err, req, res) => {
          console.error("Proxy error:", err.message);
          res.status(500).json({ error: err.message });
        },
      },
    })
  );
};