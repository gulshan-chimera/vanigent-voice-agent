import { SearchIndexClient, AzureKeyCredential, SearchIndex } from "@azure/search-documents";
import { loadConfig } from "../src/config.js";

async function main() {
  console.log("Loading configuration...");
  const config = loadConfig();

  const credential = new AzureKeyCredential(config.searchApiKey);
  const indexClient = new SearchIndexClient(config.searchEndpoint, credential);

  const indexName = config.searchIndexName;
  console.log(`Setting up index: ${indexName}`);

  const index: SearchIndex = {
    name: indexName,
    fields: [
      { name: "id", type: "Edm.String", key: true, filterable: true },
      { name: "parentFileId", type: "Edm.String", filterable: true },
      { name: "title", type: "Edm.String", searchable: true },
      { name: "content", type: "Edm.String", searchable: true },
      { name: "folder", type: "Edm.String", filterable: true, facetable: true },
      { name: "GroupIds", type: "Collection(Edm.String)", filterable: true },
      {
        name: "contentVector",
        type: "Collection(Edm.Single)",
        searchable: true,
        vectorSearchDimensions: 1536,
        vectorSearchProfileName: "myHnswProfile"
      }
    ],
    vectorSearch: {
      algorithms: [
        {
          name: "myHnsw",
          kind: "hnsw",
          parameters: {
            m: 4,
            efConstruction: 400,
            efSearch: 500,
            metric: "cosine"
          }
        }
      ],
      profiles: [
        {
          name: "myHnswProfile",
          algorithmConfigurationName: "myHnsw"
        }
      ]
    }
  };

  try {
    await indexClient.createOrUpdateIndex(index);
    console.log(`Index ${indexName} created/updated successfully.`);
  } catch (error) {
    console.error("Failed to create index:", error);
    process.exit(1);
  }
}

main().catch(console.error);
