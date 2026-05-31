---
name: context_firewall
description: "[EXPERIMENTAL] Enforces Contextual Compartmentalisation to prevent local project lore from infecting the global Laplace architecture and skills."
---
# context_firewall

**Status: [EXPERIMENTAL]**

## Purpose
The `context_firewall` skill acts as an absolute barrier between the universal, generic architecture of the Laplace identity and the local, highly specific jargon of individual projects. It ensures that any structural updates, protocol refinements, or skill modifications remain universally applicable and uncontaminated by transient project context.

## The Axiom of Compartmentalisation
1. **The Global Context is Universal**: Skills, the McCaul Protocol, and global workflows must NEVER contain project-specific nouns, concepts, or equations.
2. **The Local Context is Transient**: Project files, manuscripts, and scratchpads exist to be operated on, but their contents must not leak upward into the rules governing the operations.
3. **The Cleansing Mandate**: Before any Laplace agent writes an update to a `SKILL.md`, `mcp_config.json`, or any global configuration file, it must pass the proposed content through a strict jargon-scrubbing review.

## Sibling-project clause

The firewall covers not only upward canon-poisoning (project lore leaking into
global skills) but also lateral project-to-project poisoning. Prior or
sibling-project material - results, framings, named constructs, claims - enters a
working session only by re-derivation, never by copy. A name or framing is earned
once the current development produces what it denotes. Carrying a construct across
project boundaries without re-earning it violates this firewall in the same way
that writing it into a global skill does: it imports an unearned assumption into a
context where it has not yet been established.

In practice: if a result, concept, or framing arose in a different paper or
project, treat it as a prior reference - cite it, re-derive what you need, and
name the result only when the current work justifies the name.

## Usage Instructions
When forging or refining a skill (e.g., as `the_dreamer` or when manually prompted):
1. **Identify the Core Structural Logic**: Extract the meta-operation that needs to be encoded (e.g., "Ban complex derivations from the main text" instead of "Move the domain-specific derivation to the appendix").
2. **Scrub Local Variables**: Replace project-specific terms with abstract variables or generic functional descriptions.
3. **Verify Universal Applicability**: Ask: "Would this exact skill instruction work seamlessly for a biology paper or a software manual?" If no, the skill is too specific and must be generalized.
4. **Enforce Segregation**: If project-specific instructions are strictly necessary, relegate them to a local `project_instructions.md` within the current workspace (see [Project Scaffold](canon://templates/project_scaffold.md)), leaving the global `SKILL.md` pristine.
