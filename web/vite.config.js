import { defineConfig, loadEnv } from "vite";
import { fileURLToPath } from "node:url";

const here = fileURLToPath(new URL(".", import.meta.url));

// 개발 서버에서도 /api/g2b(나라장터 찾기 중계)가 되게 ― 배포판과 같은 ../api/g2b.mjs 를 쓴다.
// 인증키는 web/.env.local 의 DATA_GO_KR_KEY (git 에 올라가지 않는다).
function apiDev() {
  return {
    name: "api-dev",
    configureServer(server) {
      const env = loadEnv(server.config.mode, here, "");
      if (env.DATA_GO_KR_KEY && !process.env.DATA_GO_KR_KEY) process.env.DATA_GO_KR_KEY = env.DATA_GO_KR_KEY;
      server.middlewares.use("/api/g2b", async (req, res) => {
        req.url = "/api/g2b" + (req.url === "/" ? "" : req.url);
        const { default: handler } = await import("../api/g2b.mjs");
        await handler(req, res);
      });
    },
  };
}

// GitHub Pages 의 하위 경로(/저장소이름/)에서도 열리도록 상대 경로로 빌드한다.
export default defineConfig({
  base: "./",
  build: { outDir: "dist", assetsInlineLimit: 0 },
  plugins: [apiDev()],
});
