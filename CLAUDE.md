# mcp-gerard

This is Gerard McCaul's personal MCP toolkit. When working in this repo:

- Tools are in `src/mcp_gerard/`
- Shared LaTeX assets are in `src/mcp_gerard/_shared/`
- Run tests with `pytest -m unit`
- Install with `uv tool install -e .` for development

Gerard's projects live in `~/Projects/` organised by category:
  blog/, research/, simulations/, website/, vault/

Blog posts are drafted in LaTeX and published to Substack (not the website).
The website (phujck.github.io) is a static CV/landing page only.
His vault/dashboard is at ~/Projects/vault/_index.md

## Tool locations

| MCP Server | Source | Description |
|------------|--------|-------------|
| `mcp-llm` | `src/mcp_gerard/llm/tool.py` | Multi-provider LLM chat (Anthropic, OpenAI, Gemini, Groq) |
| `mcp-llm-embeddings` | `src/mcp_gerard/llm/embeddings/tool.py` | Semantic embeddings |
| `mcp-arxiv` | `src/mcp_gerard/arxiv/tool.py` | arXiv paper search + download |
| `mcp-loop` | `src/mcp_gerard/loop/tool.py` | Persistent REPL sessions via tmux |
| `mcp-code2prompt` | `src/mcp_gerard/code2prompt/tool.py` | Flatten codebases to markdown |
| `mcp-google-calendar` | `src/mcp_gerard/google_calendar/tool.py` | Calendar management |
| `mcp-word` | `src/mcp_gerard/microsoft/word/tool.py` | Word document editing |
| `mcp-blog` | `src/mcp_gerard/blog.py` | LaTeX → Substack Markdown conversion |
| `mcp-overleaf` | `src/mcp_gerard/overleaf.py` | Overleaf ↔ GitHub sync |
| `mcp-vault` | `src/mcp_gerard/vault.py` | Ideas capture, notes, project dashboard |
| `mcp-projects` | `src/mcp_gerard/projects.py` | Project creation, sync, bootstrap |

## Development

```bash
uv tool install -e .   # install in editable mode
pytest -m unit         # run unit tests only
pytest                 # run all tests
```

## Code conventions

- Always use absolute imports: `from mcp_gerard.xxx import yyy`
- Entry points expose `mcp.run` (inherited tools) or `main()` (new tools)
- New tools follow the flat-file pattern: `src/mcp_gerard/<tool>.py`
- Config from environment variables with sensible defaults
- Auto-commit in vault/projects after write operations

## Environment variables

See `.env.example` for all required variables. Key ones:
- `ANTHROPIC_API_KEY` — required for LLM tools
- `OVERLEAF_TOKEN` — required for overleaf tool
- `BLOG_DRAFTS_PATH` — defaults to `~/Projects/blog`
- `VAULT_PATH` — defaults to `~/Projects/vault`
