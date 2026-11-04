# Financial Modeling Skill

Version 0.7 brings the existing financial-modeling workflow into a three-skill repository with a public-source gap audit, an SFG visual specification and a synthetic workbook.

| Skill | Responsibility | Cowork package |
| --- | --- | --- |
| [Financial modeling](skills/financial-modeling/SKILL.md) | Workbook architecture, operating models, linked statements, valuation, financing, returns and release checks | [financial-modeling.zip](dist/financial-modeling.zip) |
| [Equity research analyst](skills/equity-research-analyst/SKILL.md) | Investment questions, company and industry evidence, competitive position and thesis risk | [equity-research-analyst.zip](dist/equity-research-analyst.zip) |
| [Valuation discipline](skills/valuation-discipline/SKILL.md) | Price-implied expectations, intrinsic and relative value, and scenario interpretation | [valuation-discipline.zip](dist/valuation-discipline.zip) |

## Reference and example

- [Source review and v0.6 gap audit](docs/SOURCE-REVIEW.md)
- [STARWAR workbook style, sheet by sheet](skills/financial-modeling/references/starwar-style.md)
- [Synthetic mini-model and reproduction instructions](examples/)
- [Skill application evaluation](docs/SKILL-EVALUATION.md)
- [Workbook style verification](docs/STYLE-VERIFICATION.md)

The modeling skill uses the existing v0.6 skill archive as its baseline. The two companion skills are standalone adaptations for repository distribution. The example is built from invented data and carries over the SFG tab names and visual conventions. Original workbook data, formulas, charts and internal training materials are excluded from the distribution.

## Cowork installation

Upload the desired ZIP through **Customize → Skills → Create skill → Upload a skill** and enable it. Each ZIP contains one top-level skill folder. Code execution must be enabled. [Claude skill packaging](https://support.claude.com/en/articles/12512198-how-to-create-custom-skills) · [Claude skill installation](https://support.claude.com/en/articles/12512180-use-skills-in-claude)

From the repository root, run `python scripts/package_skills.py` with Python 3.10 or later to rebuild and validate all three archives and their SHA-256 manifest. The synthetic workbook and its validation tools are separate from the upload packages.

## Codex installation

Clone or download this repository and copy the desired skill folder from [skills/](skills/) into your Codex skills directory. Keep its SKILL.md and supporting files together. Restart Codex if the skill does not appear.

Invoke the financial-modeling skill and describe the model, the decision it supports, and the target Excel environment.
