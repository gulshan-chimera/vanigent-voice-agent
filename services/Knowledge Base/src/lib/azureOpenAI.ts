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
 * Generates a full transcription/description of a STANDALONE image file
 * (a .png/.jpg that IS the entire document — no accompanying text layer
 * exists to supplement, unlike a PDF page). Deliberately a different
 * prompt from generateImageCaption: that one is tuned to skip content
 * already covered by separately-extracted text, and is heavily biased
 * toward returning SKIP_CAPTION — exactly wrong here, where the image
 * is the ONLY source of content and something must always be extracted
 * from it. No SKIP_CAPTION concept for this function.
 */
export async function generateImageTranscription(
  imageBase64: string,
  fileName: string,
  mimeType: string = "image/png"
): Promise<string | null> {
  const config = getConfig();
  if (!config) return null;

  const url = `${config.endpoint}/openai/deployments/${config.chatDeployment}/chat/completions?api-version=2024-06-01`;

  const prompt = `This image IS the entire content of a document named "${fileName}" — there is no separate extracted text layer for it; everything a reader needs must come from what you read in this image.

Transcribe and describe ALL informational content precisely:
- If it's a table, reproduce every row as plain text (e.g. "New Year's Day — January 1, 2025"). Do not summarize or skip rows for brevity — every row matters.
- If it's a list, form, or diagram, capture the actual content and structure, not just a general description of what kind of image it is.
- If it's a scanned or photographed document, transcribe the text as accurately as you can read it.
- Preserve exact numbers, dates, and names — do not paraphrase or round them.

Do not describe layout, colors, or formatting unless that visual structure is itself the information (e.g. a flowchart's arrows). Focus entirely on the data.

IF THERE IS NO INFORMATIONAL CONTENT AT ALL, respond with exactly this and nothing else:
${SKIP_CAPTION}

That applies to: company logos and branding, decorative graphics, banners and dividers, background textures, generic stock photography of people or offices that documents nothing specific, blank or near-blank images, and icons. These carry nothing a reader could search for, and indexing them adds noise that competes with real content in every search.

Be strict about this — a company logo is ${SKIP_CAPTION}, even though you can read the company name in it. The test is whether someone could usefully search for this image's content, not whether it contains any text at all.`;

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
              image_url: { url: `data:${mimeType};base64,${imageBase64}` },
            },
          ],
        },
      ],
      max_tokens: 1000,
      temperature: 0,
    }),
  };

  // Same retry-on-429 rationale as generateImageCaption: without it, a
  // rate limit returns null indistinguishably from a real failure, and
  // this file's ENTIRE content silently vanishes from the index.
  const MAX_ATTEMPTS = 4;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const response = await fetch(url, requestOptions);

      if (response.status === 429) {
        if (attempt === MAX_ATTEMPTS) {
          console.error(
            `[AZURE-OPENAI] Still rate limited after ${MAX_ATTEMPTS} attempts — giving up on "${fileName}".`
          );
          return null;
        }

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
        console.error(
          `[AZURE-OPENAI] Image transcription failed (${response.status}): ${errorText}`
        );
        return null;
      }

      const data = (await response.json()) as { choices: { message: { content: string } }[] };
      return data.choices[0].message.content.trim();
    } catch (error) {
      console.error(`[AZURE-OPENAI] Image transcription network error: ${(error as Error).message}`);
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
export interface SheetRecord {
  title: string;
  content: string;
}

/**
 * Turns a spreadsheet's coordinate grid into self-contained records, one
 * per logical row, each carrying its own column labels.
 *
 * Asking for JSON rather than prose is deliberate: the records become
 * search chunks directly, so we need to control the boundaries rather
 * than hoping free text splits sensibly.
 *
 * Returns null on failure so the caller can decide how to handle it.
 */
export async function generateSheetRecords(
  fileName: string,
  sheetName: string,
  cellLines: string[],
  mergedRanges: string[]
): Promise<SheetRecord[] | null> {
  const config = getConfig();
  if (!config) return null;

  const url = `${config.endpoint}/openai/deployments/${config.chatDeployment}/chat/completions?api-version=2024-06-01`;

  const gridText = cellLines.join("\n");
  const mergedText =
    mergedRanges.length > 0
      ? mergedRanges.join(", ")
      : "(none)";

  const systemPrompt = `You convert spreadsheet data into self-contained text records for a search index.

You receive every non-empty cell of one sheet, addressed by cell reference, plus the list of merged cell ranges.

YOUR TASK
Work out the sheet's structure from the coordinates, then output one record per logical row of data.

WHAT "SELF-CONTAINED" MEANS
Each record must make complete sense read on its own, with no other record for context. That means every value carries its column label inline. A record reading "Northeast, B. Smith, 42" is useless; "Region: Northeast. Sales rep: B. Smith. Accounts: 42." is correct.

STRUCTURE RULES
- Header rows are often NOT row 1. A title spanning many columns, or a merged range like A1:M1, is usually a sheet title, not a header.
- Headers can span two or three rows: a group label above, the column name below, and explanatory sub-text under that. Combine them into one meaningful label.
- Merged ranges tell you a label covers several rows or columns. A merge like A3:A4 means that label applies to both rows.
- ONE SHEET CAN CONTAIN SEVERAL SEPARATE TABLES, usually divided by blank rows. Treat each as its own table with its own headers. Never merge a row from one table with a row from another, even when they share a row label.
- Blank spacer rows between data rows are formatting, not data. Ignore them.
- Some rows are footnotes, totals, or notes rather than records. Handle them per the rules below.

NOTES AND FOOTNOTES
If a cell contains a note that qualifies a specific row (e.g. a warning that one particular class's dates may change), append it to THAT row's record. If it applies to the whole table, append it to every record in that table. Never leave a note as an orphan record with no context.

TITLES
Each record needs a short title identifying it: the row label plus which table it came from. For example "Class 5 — Recruiting & Onboarding Calendar".

OUTPUT FORMAT
Respond with ONLY valid JSON, no markdown fences, no commentary:
{"records":[{"title":"...","content":"..."}]}

Keep values exactly as they appear — do not reformat dates, correct spellings, or fix what look like errors in the source data. If a date says 2026 where 2025 seems intended, keep 2026.`;

  const userPrompt = `File: ${fileName}
Sheet: ${sheetName}
Merged ranges: ${mergedText}

Cells:
${gridText}`;

  const MAX_ATTEMPTS = 4;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
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
            { role: "user", content: userPrompt },
          ],
          max_tokens: 4000,
          temperature: 0,
          response_format: { type: "json_object" },
        }),
      });

      if (response.status === 429) {
        if (attempt === MAX_ATTEMPTS) {
          console.error(
            `[AZURE-OPENAI] Still rate limited after ${MAX_ATTEMPTS} attempts — giving up on sheet "${sheetName}".`
          );
          return null;
        }

        const retryAfterHeader = response.headers.get("retry-after");
        const baseSeconds = retryAfterHeader
          ? Number(retryAfterHeader)
          : 5 * Math.pow(2, attempt - 1);
        const waitSeconds = baseSeconds + Math.random() * 5;

        console.warn(
          `[AZURE-OPENAI] Rate limited on sheet records (attempt ${attempt}/${MAX_ATTEMPTS}) — waiting ${waitSeconds.toFixed(1)}s.`
        );
        await new Promise((resolve) => setTimeout(resolve, waitSeconds * 1000));
        continue;
      }

      if (!response.ok) {
        const errorText = await response.text();
        console.error(
          `[AZURE-OPENAI] Sheet record generation failed (${response.status}): ${errorText}`
        );
        return null;
      }

      const data = (await response.json()) as {
        choices: { message: { content: string }; finish_reason?: string }[];
      };

      const choice = data.choices[0];

      // A truncated response yields invalid JSON. Surface it rather than
      // letting the parse failure look like a model error.
      if (choice.finish_reason === "length") {
        console.error(
          `[AZURE-OPENAI] Sheet "${sheetName}" response hit the token limit — records incomplete.`
        );
        return null;
      }

      const parsed = JSON.parse(choice.message.content) as { records?: SheetRecord[] };

      if (!Array.isArray(parsed.records) || parsed.records.length === 0) {
        console.warn(`[AZURE-OPENAI] Sheet "${sheetName}" produced no records.`);
        return null;
      }

      // Drop anything malformed rather than indexing a chunk with no content.
      const valid = parsed.records.filter(
        (r) => typeof r?.title === "string" && typeof r?.content === "string" && r.content.trim().length > 0
      );

      if (valid.length < parsed.records.length) {
        console.warn(
          `[AZURE-OPENAI] Sheet "${sheetName}": dropped ${parsed.records.length - valid.length} malformed record(s).`
        );
      }

      console.log(`[AZURE-OPENAI] Sheet "${sheetName}" produced ${valid.length} record(s).`);
      return valid.length > 0 ? valid : null;
    } catch (error) {
      console.error(
        `[AZURE-OPENAI] Sheet record generation error for "${sheetName}": ${(error as Error).message}`
      );
      return null;
    }
  }

  return null;
}