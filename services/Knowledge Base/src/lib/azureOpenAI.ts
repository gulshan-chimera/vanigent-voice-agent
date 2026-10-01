// src/lib/azureOpenAI.ts
//
// Shared client helpers for calling Azure OpenAI — embeddings now,
// chat completion (for answer generation + image captioning) later.
import { teamNameForLibrary } from "./confidentiality";

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
export interface AnswerContext {
  fileName: string;
  content: string;
  restricted: boolean;
  library: string;
  score?: number;
}

export async function generateAnswer(
  question: string,
  contextChunks: AnswerContext[]
): Promise<string | null> {
  const config = getConfig();
  if (!config) return null;

  const url = `${config.endpoint}/openai/deployments/${config.chatDeployment}/chat/completions?api-version=2024-06-01`;

  const contextText = contextChunks
    .map((chunk, i) => `[Source ${i + 1}: ${chunk.fileName} | Library: ${chunk.library}]\n${chunk.content}`)
    .join("\n\n");

  // Which teams own the restricted material in this context. Named
  // explicitly so the caller is sent somewhere specific rather than
  // being told vaguely to "contact the relevant department".
  const restrictedTeams = [
    ...new Set(
      contextChunks
        .filter((c) => c.restricted)
        .map((c) => teamNameForLibrary(c.library))
    ),
  ];

  const hasRestricted = restrictedTeams.length > 0;
  const referralTeam = restrictedTeams[0] ?? "Sales Operations";

  const sensitivityBlock = hasRestricted
    ? `
CONFIDENTIAL MATERIAL IS PRESENT IN THE CONTEXT BELOW.

Some of the sources are internally confidential documents. You must judge, for this specific question, whether the caller is asking for sensitive data or for something that can be shared freely.

TREAT AS SENSITIVE — do not state these, refer instead:
- Any individual's pay, bonus, target earnings, commission, or incentive amount
- Bonus targets, payout figures, or earnings potential for any role
- Payout percentages, multipliers, attainment rates, or accelerators tied to compensation
- Performance thresholds that determine what somebody is paid
- Any dollar figure taken from a compensation or incentive plan

CAN BE SHARED FREELY — answer these normally:
- How a process works, who is eligible, when something happens
- Timing, deadlines, and submission windows
- Job expectations and responsibilities not tied to a payout figure
- Where to go, who to contact, what form to use
- Anything from a source that is not a confidential document

Judge by what the number MEANS, not by whether it is a number. "Four percent employer match" from a benefits FAQ is shareable. "One hundred percent of quarterly target" from an incentive plan is not.

WHEN THE QUESTION ASKS FOR SENSITIVE DATA, reply with only this, and nothing else:
"That's not something I can share over the phone. Please contact your ${referralTeam} team for that information."

Do not partially answer, do not hint at a range, do not confirm or deny a figure the caller suggests. If part of the question is shareable and part is not, answer the shareable part and refer the rest.
`
    : "";

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

DEPARTMENT IDENTIFICATION:
Each source has a "Library" label (e.g. "HR KB", "IT KB", "Sales Operations KB"). Use this to identify which department owns the information.
- If the caller asks "who should I contact?" or "whom should I reach out to?", tell them the department based on the library the answer came from. For example: "You would need to reach out to the HR team for that" or "That would be handled by the IT department."
- If chunks come from multiple libraries, mention all relevant departments.
- Do NOT say the library name itself — translate it to a natural department name (e.g. "HR KB" becomes "the HR team", "IT KB" becomes "the IT department", "Sales Operations KB" becomes "the Sales Operations team").
${sensitivityBlock}
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
        max_tokens: 150,
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

export async function generateSheetRecords(
  fileName: string,
  sheetName: string,
  cellLines: string[],
  mergedRanges: string[]
): Promise<SheetRecord[] | null> {
  const config = getConfig();
  if (!config) return null;

  // Group cell lines by row so a batch never splits a row across calls.
  // cellLines look like "C29: ...", so the leading number is the row.
  const rowOf = (line: string): number => {
    const m = line.match(/^[A-Z]+(\d+):/);
    return m ? Number(m[1]) : 0;
  };

  const byRow = new Map<number, string[]>();
  for (const line of cellLines) {
    const r = rowOf(line);
    const list = byRow.get(r) ?? [];
    list.push(line);
    byRow.set(r, list);
  }
  const rows = [...byRow.keys()].sort((a, b) => a - b);

  // Rows per batch. Small enough that even wide, wordy rows can't approach
  // the output token ceiling. The FAQ that lost data had ~63 rows; at 15
  // per batch that's 5 calls, each producing ~13 short records.
  const ROWS_PER_BATCH = 15;

  const batches: string[][] = [];
  for (let i = 0; i < rows.length; i += ROWS_PER_BATCH) {
    const batchRows = rows.slice(i, i + ROWS_PER_BATCH);
    batches.push(batchRows.flatMap((r) => byRow.get(r)!));
  }

  console.log(
    `[AZURE-OPENAI] Sheet "${sheetName}": ${rows.length} row(s) in ${batches.length} batch(es).`
  );

  const allRecords: SheetRecord[] = [];

  for (let b = 0; b < batches.length; b++) {
    const records = await generateRecordsForCells(
      config,
      fileName,
      sheetName,
      batches[b],
      mergedRanges,
      b + 1,
      batches.length
    );

    // A failed batch means that slice of rows is lost. Fail the whole
    // sheet rather than silently index a partial one — the caller can
    // retry, and a half-indexed sheet is the bug we're fixing.
    if (records === null) {
      console.error(
        `[AZURE-OPENAI] Sheet "${sheetName}" batch ${b + 1}/${batches.length} failed — abandoning sheet.`
      );
      return null;
    }

    allRecords.push(...records);
  }

  console.log(`[AZURE-OPENAI] Sheet "${sheetName}" produced ${allRecords.length} record(s) total.`);
  return allRecords.length > 0 ? allRecords : null;
}

async function generateRecordsForCells(
  config: NonNullable<ReturnType<typeof getConfig>>,
  fileName: string,
  sheetName: string,
  cellLines: string[],
  mergedRanges: string[],
  batchNum: number,
  batchTotal: number
): Promise<SheetRecord[] | null> {
  const url = `${config.endpoint}/openai/deployments/${config.chatDeployment}/chat/completions?api-version=2024-06-01`;

  const gridText = cellLines.join("\n");
  const mergedText = mergedRanges.length > 0 ? mergedRanges.join(", ") : "(none)";

  const systemPrompt = `You convert spreadsheet data into self-contained text records for a search index.

You receive some of the non-empty cells of one sheet, addressed by cell reference, plus the list of merged cell ranges for the whole sheet. This may be one batch of a larger sheet; produce records only for the rows present in the cells given.

YOUR TASK
Work out the structure from the coordinates, then output one record per logical row of data.

WHAT "SELF-CONTAINED" MEANS
Each record must make complete sense read on its own. Every value carries its column label inline. A record reading "Northeast, B. Smith, 42" is useless; "Region: Northeast. Sales rep: B. Smith. Accounts: 42." is correct.

STRUCTURE RULES
- Header rows are often NOT row 1. A title spanning many columns, or a merged range like A1:M1, is usually a sheet title, not a header.
- Headers can span two or three rows: combine them into one meaningful label.
- Merged ranges tell you a label covers several rows or columns.
- The header row may be OUTSIDE this batch. Infer the columns from the merged ranges and the cell contents; use the column position to label values consistently.
- Blank spacer rows are formatting, not data. Ignore them.

NOTES AND FOOTNOTES
If a cell qualifies a specific row, append it to that row's record. Never leave a note as an orphan record.

TITLES
Each record needs a short title: the row's key value plus what kind of row it is, e.g. "How is CPD calculated? — Sales Ops FAQ".

OUTPUT FORMAT
Respond with ONLY valid JSON, no markdown fences:
{"records":[{"title":"...","content":"..."}]}

Keep values exactly as they appear — do not reformat or correct anything.`;

  const userPrompt = `File: ${fileName}
Sheet: ${sheetName} (batch ${batchNum} of ${batchTotal})
Merged ranges: ${mergedText}

Cells:
${gridText}`;

  const MAX_ATTEMPTS = 4;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json", "api-key": config.apiKey },
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
          console.error(`[AZURE-OPENAI] Rate limited out on "${sheetName}" batch ${batchNum}.`);
          return null;
        }
        const retryAfterHeader = response.headers.get("retry-after");
        const baseSeconds = retryAfterHeader ? Number(retryAfterHeader) : 5 * Math.pow(2, attempt - 1);
        const waitSeconds = baseSeconds + Math.random() * 5;
        console.warn(
          `[AZURE-OPENAI] Rate limited "${sheetName}" batch ${batchNum} (attempt ${attempt}) — waiting ${waitSeconds.toFixed(1)}s.`
        );
        await new Promise((r) => setTimeout(r, waitSeconds * 1000));
        continue;
      }

      if (!response.ok) {
        console.error(`[AZURE-OPENAI] "${sheetName}" batch ${batchNum} failed (${response.status}): ${await response.text()}`);
        return null;
      }

      const data = (await response.json()) as {
        choices: { message: { content: string }; finish_reason?: string }[];
      };
      const choice = data.choices[0];

      // Still guard truncation. With 15 rows per batch this should never
      // fire, but if a batch ever does hit the limit we must NOT keep a
      // partial result — that's the silent-loss bug.
      if (choice.finish_reason === "length") {
        console.error(
          `[AZURE-OPENAI] "${sheetName}" batch ${batchNum} hit the token limit — reduce ROWS_PER_BATCH.`
        );
        return null;
      }

      const parsed = JSON.parse(choice.message.content) as { records?: SheetRecord[] };
      if (!Array.isArray(parsed.records)) return null;

      const valid = parsed.records.filter(
        (r) => typeof r?.title === "string" && typeof r?.content === "string" && r.content.trim().length > 0
      );

      return valid;
    } catch (error) {
      console.error(`[AZURE-OPENAI] "${sheetName}" batch ${batchNum} error: ${(error as Error).message}`);
      return null;
    }
  }

  return null;
}