export interface WebSearchResult {
  title: string;
  url: string;
  excerpt: string;
}

const DUCKDUCKGO_ENDPOINT = "https://api.duckduckgo.com/?q={query}&format=json&no_html=1";
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

/** Search the web with DuckDuckGo Instant Answers; no API key required. */
export async function webSearch(
  query: string,
  options: { signal?: AbortSignal } = {},
): Promise<WebSearchResult[]> {
  const trimmed = query.trim();
  if (!trimmed) throw new Error("Search query is empty");
  if (trimmed.length > 500) throw new Error("Search query is too long (500 characters max)");
  const url = DUCKDUCKGO_ENDPOINT.replace("{query}", encodeURIComponent(trimmed));
  const timeout = AbortSignal.timeout(FETCH_TIMEOUT_MS);
  const signal = options.signal ? AbortSignal.any([options.signal, timeout]) : timeout;
  let response: Response;
  try {
    response = await fetch(url, {
      signal,
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
      title: data.Heading || trimmed,
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
    if (results.length >= MAX_RESULTS) break;
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
