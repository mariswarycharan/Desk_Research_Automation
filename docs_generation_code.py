import os
from ollama import Client

client = Client(
    host="https://ollama.com",
    headers={'Authorization': 'Bearer ' + 'dd456319486541c5bfd8dfd001136b32.krWKOjU-1cHaHoKQr0iQDe0r'}
)

research_content = """
This file contains structured Excel data converted into raw text format.

The following columns are available in the Excel file:

- S NO
- Source Title
- Name of the Source
- Source  Type (Report, News, Policy, Market Data)
- Year of Publication
- Period of the Data Used
- Description
- Access Type
- Source Link
- Geography Focus (France/Italy/Germany/Spain/UK)
- Companies Mentioned
(Mindray /Snibe)

- Any other Additional Companies mentioned
- Access & Reimbursement
(Yes/No)
- Justification reference (exact content from the source)
- Explanation
- Budget Spend
(Yes/No)
- Justification reference (exact content from the source).1
- Explanation.1
- Regulatory 
(Yes/No)
- Justification reference (exact content from the source).2
- Explanation.2
- Customer Receptivity
(Yes/No)
- Justification reference (exact content from the source).3
- Explanation.3
- Others
- Justification reference (exact content from the source).4
- Explanation.4
- Summary
- Source Evidence

Below is the row-wise data in the same order:


========== Row 1 ==========
S NO: 1
Source Title: 2025 Blue Book on the Current Status and Trends of Global Expansion of Chinese Medical Devices
Name of the Source: Frost & Sullivan Consulting China
Source  Type (Report, News, Policy, Market Data): Consulting report
Year of Publication: 2025
Period of the Data Used: 2025
Description: The "2025 Blue Book on the Current Status and Trends of Global Expansion of Chinese Medical Devices" by Frost & Sullivan provides a comprehensive analysis of the Chinese medical device industry's drive for global expansion, detailing the internal factors (such as intense domestic competition from volume-based procurement (VBP) and growing R&D capacity) and external factors (such as the vast global market and the price competitiveness of Chinese devices). The report outlines key strategies for going global, including distribution, direct sales, building overseas capacity, OEM/ODM, mergers and acquisitions, and cross-border e-commerce. Despite the significant growth in exports, Chinese companies face challenges such as complex international regulatory and certification barriers (FDA PMA, EU CE-MDR), technology and intellectual property risks, and difficulties in localization. The document projects future trends, including increased R&D investment, continuous alignment with international standards, and diversification of global branding strategies.
Access Type: Open
Source Link: https://drive.google.com/file/d/1zt3vNjOlugypsvPkywduyWNRPGTHrsoj/view?usp=drive_link
Geography Focus (France/Italy/Germany/Spain/UK): France, Germany, United Kingdom, Netherlands, Belgium, Switzerland, Greece, Hungary, Turkey, Norway, Iceland, Liechtenstein, Russia
Companies Mentioned
(Mindray /Snibe)
: Both
Any other Additional Companies mentioned: Andon Health, United Imaging, Edan Instruments, Wego, Yuwell, Lepu Medical, Autobio, Fapon Biotech, Vazyme, BGI, Novogene, Jincheng Pharma, Orient Gene, Hotgen Biotech, All Test, HyTest (acquired by Mindray), DiaSys (acquired by Mindray).
Access & Reimbursement
(Yes/No): Yes
Justification reference (exact content from the source): • "EU member states commonly adopt a pricing system driven by public health insurance. National or regional agencies engage in strategic negotiations with medical device manufacturers to balance three key goals within the healthcare budget: patient affordability, reasonable company profits, and regulatory compliance."
• "This is supported by a tiered reimbursement structure, by using a DRG-based prospective payment system at the base level, and additional reimbursement mechanisms for expensive and innovative technologies at the supplementary level."
• "In Europe, where healthcare payments are mainly government-funded, companies need to prioritize public tender rules and long-term cost control."
Explanation: For IVD suppliers, Europe is not a monolithic market but a collection of government-funded systems that utilize Diagnosis-Related Groups (DRGs). Success requires navigating complex public tenders where cost-control is paramount, yet mechanisms exist for premium reimbursement of innovative diagnostics if clinical value is proven.
Budget Spend
(Yes/No): Yes
Justification reference (exact content from the source).1: • "1.65 trillion Euros Healthcare Expenditure... 3,685 Euros Healthcare Expenditure per Capita."
Explanation.1: The European market represents a massive, high-value opportunity with significant per capita spending. The budget structure is heavily weighted toward government and mandatory insurance funding (over 60% combined), meaning IVD suppliers must target public sector procurement cycles rather than private out-of-pocket spending.
Regulatory 
(Yes/No): Yes
Justification reference (exact content from the source).2: • "The EU enforces a unified and strict regulatory system for medical devices, centered on the Medical Device Regulation (MDR EU 2017/745) and the In Vitro Diagnostic Medical Device Regulation (IVDR EU 2017/746)."
• "Since the EU MDR took effect, the average CE certification process has lengthened by 6 to 12 months, and more than 30% of small and medium-sized enterprises have been forced out of the market."
• "In vitro diagnostic devices follow a separate classification under the IVDR. High-risk products require certification by Notified Bodies, while low-risk products may be self-declared by the manufacturer."
• "The MDR and IVDR transition periods have been extended to 2027-2028, requiring companies to maintain strict compliance throughout the product lifecycle."
• "The EU has strict environmental requirements for raw materials and components, such as the REACH regulation, forcing companies to adjust their R&D and design approaches to fit green supply chain management."
Explanation.2: The transition to IVDR represents a significant non-tariff barrier, lengthening certification timelines and forcing market consolidation. Chinese IVD companies face increased costs and technical hurdles (Notified Body scrutiny) compared to the previous directive era. Environmental regulations (REACH) add another layer of compliance complexity for reagents and instruments.
Customer Receptivity
(Yes/No): Yes
Justification reference (exact content from the source).3: • "Chinese medical devices face low brand recognition in overseas markets. In developed countries such as those in Europe and North America, there is price sensitivity toward low- to mid-end devices, but a lack of brand trust in high-end products from Chinese manufacturers."
• "Currently, Chinese IVD companies face multiple challenges in global expansion, including competition from international brands... In comparison, the overseas IVD market is dominated by international giants like Roche and Johnson & Johnson, which have strong brands and advanced technologies..."
• "Chinese medical device industry holds a core competitive advantage in cost-effectiveness... Chinese produced alternatives also show strong price competitiveness... This 'performance-to-cost' differentiated competition strategy allows Chinese companies... [to create] a substitution effect in international markets..."
Explanation.3: While cost-effectiveness drives adoption in low-to-mid-tier segments, a "trust deficit" hampers adoption of high-end IVD solutions in Europe. The market remains dominated by entrenched Western incumbents (Roche, etc.), making it difficult for Chinese suppliers to displace existing workflows without significant local Key Opinion Leader (KOL) endorsement or price incentives.
Others: Yes
Justification reference (exact content from the source).4: • M&A Strategy: "Mindray Medical... in 2023, it acquired Germany-based DiaSys, a company specializing in in vitro diagnostics." "On September 22, 2021, Mindray completed the acquisition of 100% equity in Finland's HyTest Invest Oy... The acquisition fills many gaps in China's upstream IVD raw materials sector."
• Market Entry Strategy: "Some Eastern European countries like Hungary and Greece... are often used by companies as gateways into the EU market. For some Class III medical devices, companies prefer certification in countries like Germany and the Netherlands, where regulations are stricter but the markets are larger, to enhance brand credibility."
• Supply Chain Risks: "Recently, the EU has strengthened supply chain localization policies... which indirectly affects the stability of the medical device supply chain and increases uncertainty for Chinese exporters."
Explanation.4: Strategic acquisitions (e.g., Mindray acquiring HyTest and DiaSys) are critical for Chinese firms to secure upstream raw materials and localized European distribution channels. Geopolitically, the EU's push for supply chain localization poses a risk to pure importers. Smart entry strategies involve utilizing specific gateway countries (Hungary/Greece) for initial access or leveraging strict German certification to build brand credibility.

Summary: The European IVD market presents a high-value but high-barrier opportunity for Chinese suppliers like Mindray, Snibe, and Autobio, driven by substantial government-backed healthcare spending. Success relies heavily on navigating the complex transition to IVDR, which has significantly raised compliance costs and timelines, favoring large, capitalized firms over smaller entrants. While Chinese suppliers hold a distinct advantage in cost-effectiveness, they face a critical "trust deficit" in the high-end segment dominated by incumbents like Roche, necessitating a strategy of acquiring local European entities (e.g., HyTest, DiaSys) to secure supply chains and credibility. Reimbursement is tightly controlled through public tenders and DRG systems, requiring suppliers to balance affordability with strict regulatory adherence. Companies are advised to utilize strategic entry points—using Eastern Europe for access or Germany for credibility—while mitigating risks associated with the EU's increasing focus on supply chain localization and environmental standards. Ultimately, the market is shifting from simple export economics to a requirement for deep localization in compliance, service, and manufacturing.

Source Evidence: https://docs.google.com/document/d/1KWqiQRt4NvZuv2_3GC2UQWi7agGABNUshk7Pk_lRDTM/edit?tab=t.0

"""

docs_prompt = f"""
Role
Act as a Senior Strategic Analyst for global diagnostics. Extract evidence-based insights revealing external forces shaping the future of the IVD market.

Objective
Evaluate the expansion potential of Chinese IVD suppliers in Europe with focus on access and reimbursement, budget structures, regulatory requirements, and customer receptivity.

INPUT RESEARCH DATA:
{research_content}


I have provided the research content above. Now, I need to convert that research content into a properly structured, professional document. Based on all the instructions I have given you, please convert it into a final well-structured professional document.

Task Instructions
Extract as many direct quotes as possible (no paraphrasing).
Briefly summarize why each piece of evidence matters for the IVD market.
Flag any assumptions, contradictions, or gaps you observe in the source.

Categories


Geography (Europe only)
Indicate whether the document mentions France, Italy, Germany, Spain, the UK, or any other European countries. If none, write “None mentioned.” (No quotes required.)


Companies Mentioned
First state whether Mindray and Snibe are mentioned or not mentioned.
Then list any other IVD companies explicitly named in the document.
No quotes required.
.
For each sub-heading below:
Answer Yes/No. If Yes, list all direct quotes as points first, then provide brief Explanations on why it is included.
Sub-headings:
Access & Reimbursement
Budget Spend
Regulatory
Customer Receptivity / Adoption
Other Relevant Insights (anything important for IVD strategy not fitting above, e.g., disease burden, lab structure, tenders, competition, etc.)


Final Summary (6–10 sentences)
Provide a concise evidence-based summary of themes, geographies, companies, access, reimbursement, regulation, adoption, risks, opportunities, and gaps.

"""


messages = [
  {
    'role': 'user',
    'content': docs_prompt,
  },
]

r = client.chat('gpt-oss:120b', messages=messages)

print(r['message']['content'], end='', flush=True)

