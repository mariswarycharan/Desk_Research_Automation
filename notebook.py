import asyncio
from notebooklm import NotebookLMClient


prompt = """

You are a Senior Research Intelligence Analyst. Your task is to extract structured evidence from a source document for the following research context:

- Research Domain  : In-Vitro Diagnostics (IVD)
- Research Topic   : European expansion potential of Chinese IVD suppliers
- Focus Entities   : Mindray, Snibe

INPUT SOURCE CONTENT: i have attached source links.


Your extraction must answer these five analytical pillars where applicable:
1. Access_Reimbursement:
Extract information related to market access, reimbursement policies, insurance coverage, public or private payer support, pricing approvals, and patient affordability. Include evidence on funding pathways, inclusion in healthcare schemes, tender access, and barriers to diagnosis adoption due to cost or reimbursement limitations.


2. Budget_Spend:
Extract information related to healthcare spending, diagnostic budgets, procurement investments, hospital or laboratory expenditure, government allocations, and funding trends or costs , money, Include data on purchasing capacity, capital investment, cost pressures, and budget priorities influencing IVD adoption.


3. Regulatory:
Extract information related to regulatory approvals, compliance requirements, certification pathways, import/export rules, quality standards, and market authorization processes. Include updates on IVDR, CE marking, local regulations, policy changes, and barriers impacting market entry.


4. Customer_Receptivity:
Extract information related to customer acceptance, clinician adoption, laboratory demand, distributor interest, brand perception, purchasing behavior, and stakeholder readiness. Include evidence on preferences for new technologies, trust in suppliers, unmet needs, and willingness to switch or adopt Chinese IVD products.


5. Any other relevant insights


---

IMPORTANT INSTRUCTIONS:

1. Do not be overly strict while extracting relevant information for each topic based on its description. If any content in the source is even slightly relevant to a topic, include it.


2. Understand the context and meaning of each sentence carefully before extracting the information. Capture what the content is actually trying to convey.


3. If multiple sentences or pieces of information are relevant to a particular topic, include all of them. Present them clearly as bullet points such as Point 1, Point 2, etc.


4. Ensure no relevant information is missed while extracting the data.



OUTPUT RULES:
- Return ONLY a valid JSON object. No markdown, no preamble.
- For every pillar: include Status (Yes/No), the EXACT first sentence of the supporting statement from the source, and a strategic Explanation.
- Include a "Summary" of 5-8 sentences synthesizing the full source relevance to the research topic.

JSON STRUCTURE:
{
"Source Title": "Give suitable title of the source document or webpage",
"Name of the Source": "Give suitable name of the publication or website",
"Source Type": "Give suitable Report, News, Policy, Market Data, etc.",
"Year of Publication": "give me Year if available",
"Period of the Data Used": "give me e.g. 2018-2023, 2024 or 'Not specified'",
"Description": "Brief short description of the source content",
"Access Type": "Free, Subscription, Paywall, Open Access, etc.",
"Source Link": "URL if applicable",
"Geography Focus": "France/Italy/Germany/Spain/UK/Global/etc.",
"Companies Mentioned (Focus)": "Comma-separated list of focus entities mentioned in the source",
"Other Entities Mentioned": "Comma-separated list of any other notable organizations/products mentioned",

"Entities_Found": "Comma-separated focus entities found (or 'None')",
"Other_Entities": "Other notable organizations/products mentioned",

"Access_Reimbursement": {
"Status": "Yes or No",
"Exact_Opening_Sentence": "First sentence of the verbatim evidence",
"Summarized_Evidence": "Paraphrased synthesis of the full evidence block",
"Explanation": "Strategic implication for the research topic"
},

"Budget_Spend": {
"Status": "Yes or No",
"Exact_Opening_Sentence": "...",
"Summarized_Evidence": "...",
"Explanation": "..."
},

"Regulatory": {
"Status": "Yes or No",
"Exact_Opening_Sentence": "...",
"Summarized_Evidence": "...",
"Explanation": "..."
},

"Customer_Receptivity": {
"Status": "Yes or No",
"Exact_Opening_Sentence": "...",
"Summarized_Evidence": "...",
"Explanation": "..."
},

"Others": {
"Exact_Opening_Sentence": "...",
"Summarized_Evidence": "...",
"Explanation": "..."
},

"Summary": "4-5 sentence synthesis of source relevance to the research."
}
"""

async def main():
    async with await NotebookLMClient.from_storage() as client:
        # Create notebook and add sources
        nb = await client.notebooks.create("Research_new")
        await client.sources.add_url(nb.id, "https://healthcare-in-europe.com/en/news/snibe-makes-an-entry-at-euromedlab.html#:~:text=%E2%80%98Our%20exports%20to%20Europe%20focus,It%E2%80%99s%20perfectly%C2%A0suited%2C%20for%20example%2C%20for", wait=True)

        # Chat with your sources
        result = await client.chat.ask(nb.id, prompt)
        print(result.answer)

        # delete notebook
        # await client.notebooks.delete(nb.id)


asyncio.run(main())
