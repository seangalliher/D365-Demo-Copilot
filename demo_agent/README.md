# D365 Demo Copilot

**AI-powered live demonstration agent for Dynamics 365 Project Operations**

The Demo Copilot takes customer requests in natural language, generates structured demo plans, and executes them in a live D365 browser session with visual overlays — spotlights, captions, business value callouts, and progress tracking.

## Key Features

| Feature | Description |
|---------|-------------|
| **Dynamic Demo Planning** | Describe what you want to demo and the AI generates a structured plan with sections and steps |
| **Tell-Show-Tell Pattern** | Every step follows presentation best practices: explain → demonstrate → summarize |
| **Visual Spotlight** | Dims the page and highlights the current element with a glowing ring |
| **Caption Overlays** | Movie-subtitle-style text at the bottom of the screen with typewriter animation |
| **Business Value Cards** | Callout cards with quantified metrics (e.g., "75% faster approval cycles") |
| **Progress Tracking** | Step counter, progress bar, and dot indicators |
| **Pause / Resume** | Pause the demo at any time — press Space or say "continue" to resume |
| **Click Ripple Effects** | Visual feedback when the agent clicks elements |
| **Title Slides** | Professional opening and closing slides |
| **Interactive Control** | Modify the plan on the fly, skip steps, or ask for a new demo |

## Architecture

```
demo_agent/
├── main.py                      # Interactive CLI entry point
├── config.py                    # Configuration (D365 URL, LLM settings)
├── requirements.txt
├── models/
│   └── demo_plan.py             # Pydantic models for demo plans & steps
├── agent/
│   ├── planner.py               # LLM-powered demo plan generator
│   ├── executor.py              # Tell-Show-Tell demo orchestrator
│   ├── narrator.py              # Dynamic narration generation
│   └── state.py                 # Pause/resume/skip state machine
├── browser/
│   ├── controller.py            # Playwright browser controller
│   ├── overlay_manager.py       # Injects/controls visual overlays
│   └── d365_pages.py            # D365-specific navigation helpers
├── overlay/
│   ├── demo-overlay.js          # Spotlight, captions, callouts, progress
│   └── demo-overlay.css         # All visual styling
├── prompts/
│   ├── planner.md               # System prompt for demo planning
│   └── narrator.md              # System prompt for narration
└── plans/
    └── sample_time_entry.json   # Example demo plan
```

## How It Works

### 1. Customer Request → Demo Plan

The customer says something like:
> "Show me how time entry and approval works in Project Operations"

The **Demo Planner** (LLM-powered) generates a structured `DemoPlan` with:
- **Sections** mapped to BPC process areas
- **Steps** with Tell-Show-Tell narration
- **Browser actions** (navigate, click, fill, spotlight)
- **Business value highlights** with quantified metrics

### 2. Demo Plan → Live Execution

The **Demo Executor** drives a Playwright browser through D365:

```
┌─────────────────────────────────────────────┐
│  Title Slide: "Project Time Tracking"       │
├─────────────────────────────────────────────┤
│  For each section:                          │
│    ├── Section transition narration         │
│    └── For each step:                       │
│         ├── TELL: Caption explaining        │
│         ├── SHOW: Browser actions +         │
│         │         spotlight/tooltips        │
│         ├── TELL: Summary caption           │
│         └── VALUE: Business value card      │
├─────────────────────────────────────────────┤
│  Closing Slide: Summary + Q&A              │
└─────────────────────────────────────────────┘
```

### 3. Visual Overlay System

A JavaScript overlay engine is injected into the D365 page, providing:

- **Spotlight**: SVG mask dims everything except the target element + glowing ring
- **Captions**: Bottom-of-screen subtitle bar with phase badges (TELL / SHOW / BUSINESS VALUE)
- **Value Cards**: Floating cards with metrics, positioned contextually
- **Progress**: Top progress bar + step dots in corner
- **Click Ripple**: Expanding ring animation on click targets
- **Pause Overlay**: Full-screen dimming with pause icon

## Getting Started

### Prerequisites

- Python 3.10+
- Node.js (for Playwright)
- Access to a D365 Project Operations environment
- **One** of: Azure OpenAI key, OpenAI API key, or a GitHub token (Copilot subscriber)

### Installation

```bash
cd demo_agent

# Create virtual environment
python -m venv .venv
.venv\Scripts\Activate.ps1   # Windows
# source .venv/bin/activate  # macOS/Linux

# Install dependencies
pip install -r requirements.txt

# Install Playwright browsers
playwright install chromium
```

### Configuration

```bash
cp .env.example .env
# Edit .env with your credentials (pick ONE):
#
#   Option 1 — Azure OpenAI:
#     AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com
#     AZURE_OPENAI_API_KEY=sk-...
#
#   Option 2 — GitHub Copilot (recommended if you have a Copilot subscription):
#     GITHUB_TOKEN=ghp_...
#     GITHUB_COPILOT_MODEL=gpt-4o          # optional, defaults to gpt-4o
#
#   Option 3 — OpenAI direct:
#     OPENAI_API_KEY=sk-...
```

> **GitHub Models bridge:** Set `GITHUB_TOKEN` to a GitHub PAT. The agent
> routes requests through the GitHub Models inference endpoint
> (`https://models.inference.ai.azure.com`) using the standard OpenAI SDK —
> no extra dependencies required.

### First Run

```bash
python -m demo_agent

# On first run, type "login" to authenticate to D365.
# The agent will save your auth state for future sessions.
```

### Running a Demo

```
Demo Copilot> Show me how project time tracking works

🧠 Generating demo plan...

📋 Demo Plan: Project Time Tracking in D365
┌────┬─────────────────┬───────────────────────┬─────────┬───────┐
│ #  │ Section         │ Step                  │ Actions │ Value │
├────┼─────────────────┼───────────────────────┼─────────┼───────┤
│ 1  │ Project Overview│ Navigate to Projects  │ 1       │ ✦     │
│ 2  │                 │ Open a Project        │ 2       │       │
│ 3  │ Time Entry      │ Navigate to Entries   │ 1       │       │
│ 4  │                 │ Create Time Entry     │ 3       │ ✦     │
│ 5  │                 │ Submit for Approval   │ 1       │ ✦     │
│ 6  │ Approval        │ Review & Approve      │ 1       │ ✦     │
└────┴─────────────────┴───────────────────────┴─────────┴───────┘

Start the demo? [yes/no/modify]: yes

🎬 Starting demo...
```

### Commands During Demo

| Command | Action |
|---------|--------|
| `pause` / `p` | Pause the demo |
| `resume` / `r` / `continue` | Resume |
| `quit` / `q` | Stop the demo |
| `plan` | Show current plan |
| `status` | Show execution status |

## Extending

### Adding Pre-Built Demo Plans

Add JSON files to `plans/` following the `DemoPlan` schema. Load them with:

```python
import json
from demo_agent.models import DemoPlan

with open("plans/sample_time_entry.json") as f:
    plan = DemoPlan(**json.load(f))
```

### Customizing Visuals

Edit `overlay/demo-overlay.css` to change colors, sizing, and animations. Key CSS variables:

```css
--demo-primary: #0078D4;        /* Microsoft Blue */
--demo-accent: #107C10;         /* Success Green */
--demo-accent-gold: #FFB900;    /* Business Value Gold */
```

### Custom Narration

The `Narrator` class can generate dynamic narration based on actual page content, supplementing the pre-planned text with context-aware commentary.

## Technology Stack

- **Python 3.10+** — Agent core
- **Playwright** — Browser automation
- **Pydantic** — Data validation and serialization
- **OpenAI / Azure OpenAI / GitHub Copilot** — LLM for demo planning and narration
- **Rich** — Terminal UI
- **Custom JS/CSS** — Visual overlay engine (zero dependencies)
