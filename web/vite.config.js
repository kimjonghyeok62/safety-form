import { defineConfig } from "vite";

// GitHub Pages 의 하위 경로(/저장소이름/)에서도 열리도록 상대 경로로 빌드한다.
export default defineConfig({
  base: "./",
  build: { outDir: "dist", assetsInlineLimit: 0 },
});
