import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { buildPrompt, extractSlides, isAuthorized, parseQuestion } from "./src/app.ts";

test("basic authentication", async () => {
  const token = btoa("j.guerreiro@unesp.br:senha");
  assert.equal(await isAuthorized(`Basic ${token}`, "j.guerreiro@unesp.br", "senha"), true);
  assert.equal(await isAuthorized(`Basic ${token}`, "j.guerreiro@unesp.br", "outra"), false);
  assert.equal(await isAuthorized("invalido", "j.guerreiro@unesp.br", "senha"), false);
});

test("question and prompt", () => {
  assert.equal(parseQuestion({ question: "  teste  " }), "teste");
  assert.throws(() => parseQuestion({ question: "" }));
  const prompt = buildPrompt("Pergunta?", [{ slide: 6, text: "Contexto." }]);
  assert.match(prompt, /\[Slide 6\] Contexto\./);
  assert.match(prompt, /nao ha informacao suficiente/);
});

test("reads slides from the presentation", async () => {
  const file = await readFile("public/vector-databases-presentation.pptx");
  const slides = extractSlides(file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength));
  assert.equal(slides.length, 10);
  assert.match(slides[0].text, /Bancos de Dados Vetoriais/i);
  assert.throws(() => extractSlides(new ArrayBuffer(0)));
});
