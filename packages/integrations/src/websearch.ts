export interface WebSearchResult {
  title: string;
  url: string;
  excerpt: string;
}

const DUCKDUCKGO_ENDPOINT = "https://api.duckduckgo.com/?q={query}&format=json&no_html=1";
const OLLAMA_SEARCH_ENDPOINT = "https://ollama.com/api/web_search";
const OLLAMA_CLOUD_ENDPOINT = "https://api.ollama.com/v1/web_search";
const FETCH_TIMEOUT_MS = 12000;
const MAX_RESULTS = 8;

interface DuckDuckGoResponse {
  AbstractText?: string;
  AbstractURL?: string;
  Heading?: string;
  RelatedTopics?: {
    Text?: string;
    FirstURL?: string;
    Topics?: { Text?: string; FirstURL?: string }[];
  }[];
}

interface OllamaSearchResponse {
  results?: {
    title?: string;
    url?: string;
    link?: string;
    snippet?: string;
    content?: string;
    description?: string;
    source?: string;
    published_date?: string;
  }[];
}

/** Search the web: Ollama Web Search first (when OLLAMA_API_KEY is set), DuckDuckGo fallback. */
export async function webSearch(
  query: string,
  options: { signal?: AbortSignal; maxResults?: number } = {},
): Promise<WebSearchResult[]> {
  const trimmed = query.trim();
  if (!trimmed) throw new Error("Search query is empty");
  if (trimmed.length > 500) throw new Error("Search query is too long (500 characters max)");
  const maxResults = options.maxResults ?? MAX_RESULTS;
  const ollamaKey = process.env.OLLAMA_API_KEY?.trim();
  if (ollamaKey) {
    try {
      const results = await ollamaWebSearch(trimmed, ollamaKey, maxResults, options.signal);
      if (results.length) return results;
    } catch {
      // Ollama search failed or rate-limited; fall through to the keyless provider.
    }
  }
  return duckDuckGoSearch(trimmed, maxResults, options.signal);
}

/** Ollama's official web search API (https://docs.ollama.com/capabilities/web-search). */
async function ollamaWebSearch(
  query: string,
  apiKey: string,
  maxResults: number,
  signal?: AbortSignal,
): Promise<WebSearchResult[]> {
  const timeout = AbortSignal.timeout(FETCH_TIMEOUT_MS);
  const combined = signal ? AbortSignal.any([signal, timeout]) : timeout;
  const body = JSON.stringify({ query, max_results: maxResults });
  const attempts = [OLLAMA_SEARCH_ENDPOINT, OLLAMA_CLOUD_ENDPOINT].map((endpoint) =>
    fetch(endpoint, {
      method: "POST",
      signal: combined,
      headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
      body,
    }).then((response) => ({ endpoint, response })),
  );
  const settled = await Promise.allSettled(attempts);
  for (const outcome of settled) {
    if (outcome.status !== "fulfilled" || !outcome.value.response.ok) continue;
    const data = (await outcome.value.response.json()) as OllamaSearchResponse;
    const results: WebSearchResult[] = (data.results ?? [])
      .filter((entry) => (entry.url || entry.link) && entry.title)
      .slice(0, maxResults)
      .map((entry) => ({
        title: (entry.title || "").slice(0, 160),
        url: entry.url || entry.link || "",
        excerpt: (entry.snippet || entry.content || entry.description || "").slice(0, 600),
      }));
    if (results.length) return results;
  }
  return [];
}

/** Search the web with DuckDuckGo Instant Answers; no API key required. */
async function duckDuckGoSearch(
  query: string,
  maxResults: number,
  signal?: AbortSignal,
): Promise<WebSearchResult[]> {
  const url = DUCKDUCKGO_ENDPOINT.replace("{query}", encodeURIComponent(query));
  const timeout = AbortSignal.timeout(FETCH_TIMEOUT_MS);
  const combined = signal ? AbortSignal.any([signal, timeout]) : timeout;
  let response: Response;
  try {
    response = await fetch(url, {
      signal: combined,
      headers: { "User-Agent": "OpenMuse/1.0 (personal agent)" },
    });
  } catch (error: unknown) {
    if (timeout.aborted) throw new Error("Search timed out after 12 seconds");
    throw error;
  }
  if (!response.ok) throw new Error(`Search service returned HTTP ${response.status}`);
  const data = (await response.json()) as DuckDuckGoResponse;
  const results: WebSearchResult[] = [];
  if (data.AbstractText && data.AbstractURL)
    results.push({
      title: data.Heading || query,
      url: data.AbstractURL,
      excerpt: stripHtml(data.AbstractText).slice(0, 600),
    });
  const flatten = (
    topics: DuckDuckGoResponse["RelatedTopics"],
  ): { Text?: string; FirstURL?: string }[] =>
    (topics ?? []).flatMap((topic) => (topic.FirstURL ? [topic] : flatten(topic.Topics)));
  for (const topic of flatten(data.RelatedTopics)) {
    if (!topic.FirstURL || !topic.Text) continue;
    results.push({
      title: topic.Text.split(" - ")[0].slice(0, 160),
      url: topic.FirstURL,
      excerpt: topic.Text.slice(0, 600),
    });
    if (results.length >= maxResults) break;
  }
  return results;
}

/** Strip HTML tags and collapse whitespace in a page snippet. */
function stripHtml(html: string): string {
  return html
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
