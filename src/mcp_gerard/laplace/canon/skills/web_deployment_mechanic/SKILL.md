---
name: web_deployment_mechanic
description: "[DEPRECATED] Automates the deployment of finished manuscripts to the configured web repository, including summary generation, HTML updating, and PDF copying."
---

# Web Deployment Mechanic [DEPRECATED]

This skill automates the high-friction manual task of publishing an approved manuscript to the configured web repository.

## Deployment Protocol

When directed to publish or deploy a manuscript to the website, follow these steps exactly:

1. **PDF Synchronisation**: Copy the finalised, compiled `.pdf` from the manuscript directory to the assets/docs folder of the configured web repository. Ensure the filename is lowercased and hyphenated.

2. **The Ordinary-Language Summary**: Read the final abstract and introduction of the manuscript. Draft a roughly 500-word summary in ordinary language. Include a disclaimer stating the work is experimental and should not be considered final.

3. **Nature-Style Visual Embedding**: Generate or extract one or two Nature-style visual elements (Mermaid diagrams or clean vector SVGs) representing the core conceptual bridge. Ensure these visuals obey the Child Made It rule: clean flat pastel fills, thick dark outlines, muted vermilions.

4. **HTML Injection**: Open the relevant section page in the configured web repository. Append the new entry, summary, disclaimer, and PDF link to the appropriate section, observing the `html_mechanic` (British English, spaced hyphens, Vonnegut Rule).

5. **Commit and Push**: Use the `git_mechanic` skill to stage, commit, and push the changes to the configured web repository.

## Operational Constraints
- Never deploy a manuscript that has not passed the full `laplace_release_mechanic` pipeline.
- The web deployment is the final physical manifestation of the manuscript. Treat its aesthetic quality with the same rigour as the LaTeX source.
