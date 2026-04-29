import os
from ollama import Client

client = Client(
    host="https://ollama.com",
    headers={'Authorization': 'Bearer ' + 'dd456319486541c5bfd8dfd001136b32.krWKOjU-1cHaHoKQr0iQDe0r'}
)



# =====================================================
# PURE PYTHON TOKEN COUNTER (NO EXTERNAL LIBRARIES)
# Approximate token count for DeepSeek / GPT / LLMs
# =====================================================

import re
import math


def estimate_tokens(text):
    """
    Approximate token count using Python logic only.

    Rule:
    - English words ≈ 0.75 tokens per word
    - Punctuation counted separately
    - Numbers counted separately
    - Spaces ignored
    """

    # Split words / punctuation / numbers
    parts = re.findall(r"\w+|[^\w\s]", text, re.UNICODE)

    token_count = 0

    for item in parts:

        # If number
        if item.isdigit():
            token_count += 1

        # If punctuation
        elif len(item) == 1 and not item.isalnum():
            token_count += 1

        # If word
        else:
            # Approximation:
            # 1 token for every 4 characters
            token_count += math.ceil(len(item) / 4)

    return token_count


def analyze_text(text):
    words = len(text.split())
    chars = len(text)
    tokens = estimate_tokens(text)

    print("=" * 60)
    print("PURE PYTHON TOKEN ESTIMATION")
    print("=" * 60)
    print(f"Characters : {chars}")
    print(f"Words      : {words}")
    print(f"Tokens     : {tokens}")
    print("=" * 60)


with open("txts/xlsx_to_txt_content.txt", "r", encoding="utf-8") as file:
    research_content = file.read()


# with open("row1_docs.txt", "r", encoding="utf-8") as file:
#     research_content = file.read()

# with open("row2_docs.txt", "r", encoding="utf-8") as file:
#     research_content += "\n\n" + file.read()

pptx_prompt = f"""
You are a Senior PowerPoint Automation Specialist and Python python-pptx Code Generation Expert.

Your responsibility is to generate a complete, production-ready, error-free Python script using the python-pptx library that creates a highly professional client-facing research presentation deck.

INPUT RESEARCH DATA CONTENT:
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
Generate Python code only using python-pptx that builds a complete PowerPoint deck (15 - 20 slides):

STRICT CONTENT RULES:
1. Every factual statement in slides must come only from the provided research input.
2. Do not invent facts, statistics, claims, company names, or numbers.
3. If some information is not available, intelligently omit it.
4. Use professional consulting-style language.
6. Ensure logical storytelling flow across slides.
7. Maintain continuity between slides.


CONTENT IMPACT RULES:
1. Create strong, executive-style slide titles that are catchy, meaningful, insight-led, and immediately valuable to the client. Every title should clearly reflect the slide message and create interest.
2. Slide content must be concise, impactful, and professionally written so that each point adds value, communicates insight, and creates a positive impression while reading.
3. Tables, charts, and visuals must also communicate clear business meaning, using smart labels, structured formatting, and attractive presentation that strengthens the overall client impact.

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
1. Use bolded word for title and heading text and make it in blue colour and keep font size as 16.
1. Premium consulting / strategy firm quality.
2. Clean white-space usage.
3. Strong visual hierarchy.
4. Professional fonts.
5. Use bold headings.
6. Use italic selectively.
7. Proper text alignment.
8. Use justified body text where suitable.
10. Use structured content boxes.
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

IMPORTANT: Use only data available in input research content.

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


IMPORTANT INSTRUCTIONS:
- You must want to include tabular data presentation in atleast 2 - 3 slides.
- You must want to include charts, visuals, in the atleast 2 - 3 slides , to create visual you can use matplotlib, seaborn, plotly kind of libraries and include code for it in the final code output

1. Content quality is very important for every slide. Include as much relevant content as possible while keeping it properly aligned and well-structured.

2. Avoid slides with very little content. Maintain a balanced, high level of content on each slide.

3. Ensure each slide contains essential and meaningful information relevant to its title.

4. Important points must be included in every slide so that the presentation looks professional, informative, and engaging for the reader.

5. Bullet points should be used for all content in all slides.



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

# count tokens in prompt
analyze_text(pptx_prompt)

# deepseek-v3.2:cloud
# deepseek-v4-pro:cloud

r = client.chat('deepseek-v3.1:671b-cloud', messages=messages, think='high')

final_response = r['message']['content']

with open("generated_pptx_code.py", "w", encoding="utf-8") as file:
    file.write(final_response)

