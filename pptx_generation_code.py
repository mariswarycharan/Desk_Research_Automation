import os
from ollama import Client

client = Client(
    host="https://ollama.com",
    headers={'Authorization': 'Bearer ' + 'dd456319486541c5bfd8dfd001136b32.krWKOjU-1cHaHoKQr0iQDe0r'}
)


with open("xlsx_to_txt_content.txt", "r", encoding="utf-8") as file:
    research_content = file.read()


# with open("row1_docs.txt", "r", encoding="utf-8") as file:
#     research_content = file.read()

# with open("row2_docs.txt", "r", encoding="utf-8") as file:
#     research_content += "\n\n" + file.read()

pptx_prompt = f"""
You are a Senior PowerPoint Automation Specialist and Python python-pptx Code Generation Expert.

Your responsibility is to generate a complete, production-ready, error-free Python script using the python-pptx library that creates a highly professional client-facing research presentation deck.

INPUT RESEARCH DATA:
{research_content}


INPUT DATA CONTEXT:
You will receive structured research insights converted from Excel-based desk research outputs. The original data was extracted from multiple validated public sources such as reports, websites, policy documents, market intelligence sources, company releases, journals, and news articles.

The input dataset may contain rows/records with fields such as:

- Source Title
- Name of Source
- Source Type
- Year of Publication
- Period of Data Used
- Geography Focus
- Companies Mentioned
- Other Entities Mentioned
- Description
- Summary
- Access & Reimbursement
- Budget & Spend
- Regulatory Environment
- Customer / Stakeholder Receptivity
- Other Insights
- Source Link
- Quantitative Metrics
- Market Numbers
- Growth Rates
- Strategic Insights
- Country / Region
- Segment
- Any additional research findings

You must analyze the full provided input content and convert it into a polished executive presentation.

CORE OBJECTIVE:
Generate Python code only using python-pptx that builds a complete PowerPoint deck (15 - 20 slides) using this template file:

Use the template theme, layout style, formatting logic, and branding across all generated slides.

STRICT CONTENT RULES:
1. Every factual statement in slides must come only from the provided research input.
2. Do not invent facts, statistics, claims, company names, or numbers.
3. If some information is not available, intelligently omit it.
4. Use professional consulting-style language.
5. Summarize dense data into executive-ready concise insights.
6. Ensure logical storytelling flow across slides.
7. Maintain continuity between slides.

DECK STRUCTURE:
Automatically create an optimal storyline and table of contents based on input data.

Recommended flow:

1. Title Slide
2. Executive Summary
3. Agenda / Table of Contents
4. Market Overview
5. Key Trends
6. Geography Insights
7. Competitive Landscape
8. Access & Reimbursement Insights
9. Budget & Spend Analysis
10. Regulatory Landscape
11. Customer / Stakeholder Receptivity
12. Strategic Opportunities
13. Risks / Challenges
14. Key Metrics & Charts
15. Recommendations
16. Conclusion
17. Appendix / Sources (if needed)

You may adapt structure depending on available input data.

SLIDE DESIGN RULES:
1. Premium consulting / strategy firm quality.
2. Clean white-space usage.
3. Strong visual hierarchy.
4. Professional fonts.
5. Use bold headings.
6. Use italic selectively.
7. Proper text alignment.
8. Use justified body text where suitable.
9. Consistent spacing.
10. Use structured content boxes.
11. Use smart section dividers.
12. Use visually balanced layouts.
13. No emojis.

VISUAL CONTENT RULES:
Where relevant, generate and include:

- Bar charts
- Line charts
- Pie charts
- Tables
- Comparison matrices
- KPI highlight cards
- Timelines
- Heatmaps (if feasible)
- Trend summaries

Use only data available in input.

IMAGES / LOGOS:
If relevant to companies, countries, sectors, or themes:
- Use professional logos
- Use professional icons
- Use suitable images only if necessary
- Maintain premium corporate appearance

REFERENCE CITATION RULES:
This is mandatory.

For every research-backed statement included in slides:

Example:
Market grew by 12% YoY [1]
France reimbursement reforms accelerated adoption [2]

At the bottom of each slide, include references used in that slide in single text box and one by one pointer wise:

[1] https://source-link-1.com
[2] https://source-link-2.com
[3] https://source-link-3.com

Rules:
1. Reference numbering should be slide-specific.
2. Only cite sources actually used on that slide.
3. Every sourced claim should have citation tag.
4. General non-factual headings do not need citations.

TABLE RULES:
If tabular presentation improves clarity:
- Create clean tables
- Styled headers
- Proper row spacing
- Readable font size
- Aligned columns

CHART RULES:
If metrics exist:
- Convert metrics into charts
- Add titles
- Add labels
- Use professional color palette
- Ensure readability


IMPORTANT:
- You must want to include tabular data presentation in atleast 2 - 3 slides.
- You must want to include charts, visuals, in the atleast 2 - 3 slides , to create visual you can use matplotlib, seaborn, plotly kind of libraries and include code for it in the final code output


CODE GENERATION RULES:
Generate only Python code.

Mandatory requirements:

1. Use python-pptx only.
3. Final script must save PPT locally.
4. Code must run without errors.
5. No comments anywhere in code.
6. No explanations before code.
7. No explanations after code.
8. No markdown.
9. No placeholders.
10. Fully executable script.
11. Include imports.
12. Handle text overflow smartly.
13. Use reusable helper functions.
14. Maintain consistent styling.
15. Ensure all slides render properly.

TEXT QUALITY RULES:
1. Executive tone
2. Concise
3. Insightful
4. Strategic
5. Client-ready
6. Grammatically strong
7. High readability

SLIDE COUNT:
Generate 15-20 high-quality slides depending on content richness.

FINAL OUTPUT RULE:
Return only the complete Python code script. Nothing else.
"""

messages = [
  {
    'role': 'user',
    'content': pptx_prompt,
  },
]

r = client.chat('deepseek-v3.1:671b-cloud', messages=messages)

print(r['message']['content'], end='', flush=True)

