// src/lib/azureOpenAI.ts
//
// Shared client helpers for calling Azure OpenAI — embeddings now,
// chat completion (for answer generation + image captioning) later.

function getConfig() {
  const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const apiKey = process.env.AZURE_OPENAI_API_KEY;
  const embeddingDeployment = process.env.AZURE_OPENAI_EMBEDDING_DEPLOYMENT;
  const chatDeployment = process.env.AZURE_OPENAI_CHAT_DEPLOYMENT;

  if (!endpoint || !apiKey || !embeddingDeployment || !chatDeployment) {
    console.error("[AZURE-OPENAI] Missing one or more Azure OpenAI environment variables.");
    return null;
  }

  return { endpoint: endpoint.replace(/\/$/, ""), apiKey, embeddingDeployment, chatDeployment };
}

/**
 * Generates an embedding vector for the given text using the deployed
 * text-embedding-3-small model.
 */
export async function generateEmbedding(text: string): Promise<number[] | null> {
  const config = getConfig();
  if (!config) return null;

  const url = `${config.endpoint}/openai/deployments/${config.embeddingDeployment}/embeddings?api-version=2024-02-01`;

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": config.apiKey,
      },
      body: JSON.stringify({ input: text }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`[AZURE-OPENAI] Embedding request failed (${response.status}): ${errorText}`);
      return null;
    }

    const data = (await response.json()) as { data: { embedding: number[] }[] };
    return data.data[0].embedding;
  } catch (error) {
    console.error(`[AZURE-OPENAI] Embedding network error: ${(error as Error).message}`);
    return null;
  }
}

/**
 * Generates a caption describing the visual elements of a page image,
 * using that page's own extracted text as context — so the caption
 * connects to the surrounding instructions rather than describing the
 * image in isolation.
 */
/**
 * Generates a caption describing meaningful visual content on a page,
 * using that page's own extracted text as context.
 *
 * The prompt is deliberately strict: earlier testing showed the model
 * would always find *something* to describe (logos, section header bars,
 * footers), producing captions that added no retrievable information and
 * diluted the page's embedding with noise. It now names those elements
 * explicitly as ignorable, and requires the caption to contain
 * information not already present in the page text.
 *
 * Returns null on failure, or the exact string SKIP_CAPTION when there
 * is nothing worth describing.
 */
export const SKIP_CAPTION = "NOTHING_TO_DESCRIBE";

export async function generateImageCaption(
  imageBase64: string,
  pageText: string
): Promise<string | null> {
  const config = getConfig();
  if (!config) return null;

  const url = `${config.endpoint}/openai/deployments/${config.chatDeployment}/chat/completions?api-version=2024-06-01`;

  const prompt = `This image is a page from a company HR document. Here is the text already extracted from this same page:

"""${pageText}"""

Your job is to describe visual content that carries INFORMATION A READER WOULD OTHERWISE MISS — and nothing else.

DESCRIBE these:
- Screenshots of software interfaces, especially where an arrow, box, or highlight points at a specific button, tab, menu, or field
- Charts, graphs, or diagrams carrying data or relationships
- Flowcharts, org charts, or process illustrations
- Photographs of equipment, facilities, or procedures being demonstrated
- Any figure whose meaning is NOT already stated in the page text above

IGNORE these completely — they are page furniture, not content:
- Company logos or branding of any kind
- Section headers, coloured header bars, or title banners
- Footers, page numbers, dates, document IDs
- Text formatting: bold, italics, numbered lists, bullets, indentation
- Table borders, rules, dividers, background shading, decorative graphics
- Generic stock photography of people, offices, or meetings that illustrates nothing specific

DECISION RULE: if everything visual on this page falls into the IGNORE list, or if the visual content only repeats what the page text already says, respond with exactly this and nothing else:
${SKIP_CAPTION}

Most pages in a text document will be ${SKIP_CAPTION}. That is the expected and correct answer — do not look for something to say.

If there IS genuine visual content, describe only that. Do not mention the ignored elements. Do not transcribe body text. If a screenshot shows a specific step, state exactly what it shows and which step it belongs to.`;

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": config.apiKey,
      },
      body: JSON.stringify({
        messages: [
          {
            role: "user",
            content: [
              { type: "text", text: prompt },
              {
                type: "image_url",
                image_url: { url: `data:image/png;base64,${imageBase64}` },
              },
            ],
          },
        ],
        max_tokens: 500,
        temperature: 0,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`[AZURE-OPENAI] Caption request failed (${response.status}): ${errorText}`);
      return null;
    }

    const data = (await response.json()) as { choices: { message: { content: string } }[] };
    return data.choices[0].message.content.trim();
  } catch (error) {
    console.error(`[AZURE-OPENAI] Caption network error: ${(error as Error).message}`);
    return null;
  }
}
/**
 * Generates a natural, spoken-style answer to the caller's question,
 * grounded strictly in the provided context chunks. Returns null on
 * failure so the caller can decide how to handle it (fail closed).
 */
export async function generateAnswer(
  question: string,
  contextChunks: { fileName: string; content: string }[]
): Promise<string | null> {
  const config = getConfig();
  if (!config) return null;

  const url = `${config.endpoint}/openai/deployments/${config.chatDeployment}/chat/completions?api-version=2024-06-01`;

  const contextText = contextChunks
    .map((chunk, i) => `[Source ${i + 1}: ${chunk.fileName}]\n${chunk.content}`)
    .join("\n\n---\n\n");

  const systemPrompt = `You are answering an employee's question over a phone call, using only the internal HR documents provided below as context.

RULES:
- Answer using ONLY the information in the provided context. Never add information not present there.
- If the context does not contain the answer, say exactly: "I don't have that information right now, but I can note it for the team to follow up." Do not guess or make anything up.
- Keep the answer to 1-3 short sentences, since this will be spoken aloud on a phone call.
- Do not use any markdown, bullet points, or formatting — plain spoken sentences only.
- Spell out numbers, percentages, and dates in natural spoken form (e.g. "four percent" not "4%").
- Do not mention "the context," "the document," "Source 1," or that you were given reference material — just answer naturally, as if you already knew it.

CONTEXT:
${contextText}`;

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": config.apiKey,
      },
      body: JSON.stringify({
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: question },
        ],
        max_tokens: 300,
        temperature: 0.3,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`[AZURE-OPENAI] Answer generation failed (${response.status}): ${errorText}`);
      return null;
    }

    const data = (await response.json()) as { choices: { message: { content: string } }[] };
    return data.choices[0].message.content.trim();
  } catch (error) {
    console.error(`[AZURE-OPENAI] Answer generation network error: ${(error as Error).message}`);
    return null;
  }
}