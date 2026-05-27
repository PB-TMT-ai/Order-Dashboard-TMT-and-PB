# Project: Order-Dashboard-TMT-and-PB

## Core Principle

You operate as the decision-maker in a modular system. Your job is NOT to do everything yourself. Your job is to read instructions, pick the right tools, handle errors intelligently, and improve the system over time.

Why? 90% accuracy across 5 steps = 59% total success. Push repeatable work into tested scripts. You focus on decisions.

## System Architecture

**Blueprints (/blueprints)** - Step-by-step instructions in markdown. Goal, inputs, scripts to use, output, edge cases. Check here FIRST.

**Scripts (/scripts)** - Tested, deterministic code. Call these instead of writing from scratch.

**Workspace (/.workspace)** - Temp files. Never commit. Delete anytime.

## How You Operate

1. Check blueprints first - If one exists, follow it exactly
2. Use existing scripts - Only create new if nothing exists
3. Fail forward - Error → Fix → Test → Update blueprint → Add to LEARNINGS.md → System smarter
4. Ask before creating - Don't overwrite blueprints without asking

## Tech Stack

- React
- TypeScript
- Tailwind CSS
- Supabase
- Next.js

## Project Structure

/src (components, pages, hooks, lib, types, styles)
/scripts - Automation scripts
/blueprints - Task SOPs
/.workspace - Temp files (gitignored)

## Code Standards

- TypeScript strict, explicit return types
- Functional components only
- Props: ComponentNameProps
- No any - use unknown
- Async/await over .then()

## Error Protocol

1. Stop and read the full error
2. Isolate - which component/script failed
3. Fix and test
4. Document in LEARNINGS.md
5. Update relevant blueprint

## What NOT To Do

- Don't skip blueprint check
- Don't ignore errors and retry blindly
- Don't create files outside structure
- Don't write from scratch when blueprint exists
