import assert from "node:assert/strict";
import { test } from "node:test";
import {
  assertApiDeploymentConfig,
  type Config,
  DEFAULT_MODEL_TIMEOUT_MS,
  positiveInt,
  shadowedEnvKeys,
} from "../apps/server/src/config.ts";

const sampleConfig: Config = {
  mode: "sample",
  port: 8787,
  host: "127.0.0.1",
  publicUrl: "http://localhost:8787",
  dataDir: ".openmuse",
  agentBackend: "sample",
  googleRedirectUri: "http://localhost:8787/api/google/callback",
  allowedOrigins: ["http://localhost:8081"],
};

function liveConfig(intelligenceApiKey?: string): Config {
  return {
    ...sampleConfig,
    mode: "live",
    agentBackend: "model",
    intelligenceApiKey,
  };
}

const missingKeyMessage =
  "OpenMuse requires CPK_INTELLIGENCE_API_KEY. " +
  "Run `npx copilotkit@latest login` and `npx copilotkit@latest project select`, " +
  "then set the generated server-only key. " +
  "See https://docs.copilotkit.ai/intelligence/connect-your-runtime";

test("every API mode rejects a missing or blank Intelligence key", () => {
  for (const mode of [sampleConfig, liveConfig()]) {
    for (const key of [undefined, "", " \t\n"]) {
      assert.throws(() => assertApiDeploymentConfig({ ...mode, intelligenceApiKey: key }), {
        name: "Error",
        message: missingKeyMessage,
      });
    }
  }
});

test("every API mode accepts a non-empty Intelligence key", () => {
  for (const mode of [sampleConfig, liveConfig()]) {
    assert.doesNotThrow(() =>
      assertApiDeploymentConfig({ ...mode, intelligenceApiKey: "test-project-key-never-sent" }),
    );
  }
});

test("environment variables that override a different .env value are reported by name", () => {
  const file = { OPENAI_API_KEY: "sk-or-file", MODEL: "openai/gpt-5", PORT: "8787", EMPTY: "" };
  const env = { OPENAI_API_KEY: "sk-proj-system", MODEL: "openai/gpt-5", EMPTY: "set" };
  assert.deepEqual(shadowedEnvKeys(file, env), ["OPENAI_API_KEY", "EMPTY"]);
  assert.deepEqual(shadowedEnvKeys(file, {}), []);
});

test("positiveInt accepts a positive number and floors it", () => {
  assert.equal(positiveInt("X", 100, "2500"), 2500);
  assert.equal(positiveInt("X", 100, "2500.9"), 2500);
});

test("positiveInt falls back for missing, blank, zero, negative or non-numeric values", () => {
  assert.equal(positiveInt("X", DEFAULT_MODEL_TIMEOUT_MS, undefined), DEFAULT_MODEL_TIMEOUT_MS);
  assert.equal(positiveInt("X", 100, ""), 100);
  assert.equal(positiveInt("X", 100, "   "), 100);
  assert.equal(positiveInt("X", 100, "0"), 100);
  assert.equal(positiveInt("X", 100, "-5"), 100);
  assert.equal(positiveInt("X", 100, "abc"), 100);
  assert.equal(positiveInt("X", 100, "Infinity"), 100);
});

test("the default model timeout is longer than the old five-minute cap", () => {
  assert.ok(DEFAULT_MODEL_TIMEOUT_MS > 5 * 60 * 1000);
});
