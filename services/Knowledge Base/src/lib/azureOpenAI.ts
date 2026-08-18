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