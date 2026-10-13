import esbuild from "esbuild";
import { createPluginBundlerPresets } from "@paperclipai/plugin-sdk/bundlers";

const presets = createPluginBundlerPresets({ uiEntry: "src/ui/index.tsx" });
const contexts = await Promise.all([
  esbuild.context(presets.esbuild.worker),
  esbuild.context(presets.esbuild.manifest),
  esbuild.context(presets.esbuild.ui),
]);
try {
  await Promise.all(contexts.map((context) => context.rebuild()));
} finally {
  await Promise.all(contexts.map((context) => context.dispose()));
}
