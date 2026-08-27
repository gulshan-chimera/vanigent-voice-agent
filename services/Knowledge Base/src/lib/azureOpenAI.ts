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

  const prompt = `This image is a page from an internal company document. Here is the text already extracted from this same page:

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

  const requestOptions = {
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
  };

  // Retry on 429 with exponential backoff. Without this, a rate limit
  // returns null, which the caller cannot distinguish from "this page
  // has nothing worth describing" — so real screenshots silently vanish
  // from the index while the sync still reports success.
  const MAX_ATTEMPTS = 4;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const response = await fetch(url, requestOptions);

      if (response.status === 429) {
        if (attempt === MAX_ATTEMPTS) {
          console.error(
            `[AZURE-OPENAI] Still rate limited after ${MAX_ATTEMPTS} attempts — giving up on this page.`
          );
          return null;
        }

        // Azure tells us how long to wait; fall back to exponential
        // backoff (5s, 10s, 20s) when the header is absent.
        const retryAfterHeader = response.headers.get("retry-after");
        const baseSeconds = retryAfterHeader
          ? Number(retryAfterHeader)
          : 5 * Math.pow(2, attempt - 1);

        const waitSeconds = baseSeconds + Math.random() * 5;
        
        console.warn(
          `[AZURE-OPENAI] Rate limited (attempt ${attempt}/${MAX_ATTEMPTS}) — waiting ${waitSeconds}s.`
        );
        await new Promise((resolve) => setTimeout(resolve, waitSeconds * 1000));
        continue;
      }

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

  return null;
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

  const systemPrompt = `You are answering an employee's question over the phone, using only the internal company documents provided below.

HOW THESE DOCUMENTS ARE STRUCTURED — read this carefully:
Many are FAQ tables where ONE answer covers SEVERAL questions grouped together. When text is extracted, the answer often attaches to only the first question in the group, leaving the others looking unanswered. If you see a run of related questions with a single answer nearby, that answer applies to ALL of them. Use it.

Many answers are SIGNPOSTS rather than facts — they direct the employee to a system, a phone number, or another document instead of stating a figure or a date. A signpost IS a valid answer. Relay it. Do not treat it as missing information.

Examples of valid answers:
- Asked when benefits begin, and the document says eligibility is managed through ADP TotalSource with a number to call: tell them it's handled through ADP and give them the number.
- Asked about 401(k) eligibility, and the document says to view the 401(k) Plan Highlights: tell them that's where it's set out, and mention who administers the plan if the document says.

RULES:
- Use ONLY the information in the context below. Never add anything not present.
- Give the most useful thing the context does contain, even if it's a pointer rather than a direct fact.
- Only say "I don't have that information right now, but I can note it for the team to follow up" when the context contains nothing relevant to the question at all — not merely because it lacks the exact figure asked for.
- Keep it to one or two short sentences. This is spoken aloud.
- No markdown, no lists, no formatting.
- Spell numbers, percentages and dates in spoken form. Phone numbers digit by digit.
- Never mention the context, the documents, sources, or that you looked anything up.

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
        temperature: 0,
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