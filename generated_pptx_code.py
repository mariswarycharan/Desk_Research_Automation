from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.xmlchemy import OxmlElement
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import numpy as np
import io

def create_presentation():
    prs = Presentation()
    
    title_slide_layout = prs.slide_layouts[0]
    content_slide_layout = prs.slide_layouts[1]
    section_header_layout = prs.slide_layouts[2]
    
    def add_title_slide():
        slide = prs.slides.add_slide(title_slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        title.text = "Chinese IVD Expansion in European Markets"
        subtitle.text = "Strategic Analysis and Market Entry Assessment"
    
    def add_executive_summary():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Executive Summary: High-Value Market with Significant Barriers"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "• European IVD market represents €12.9B+ opportunity with substantial government-funded healthcare spending"
        p.font.size = Pt(14)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Chinese suppliers face complex regulatory barriers (IVDR transition) and new procurement restrictions (IPI)"
        p.font.size = Pt(14)
        
        p = tf.add_paragraph()
        p.text = "• Leading Chinese players (Mindray, Snibe, Autobio) demonstrate regulatory compliance and financial strength"
        p.font.size = Pt(14)
        
        p = tf.add_paragraph()
        p.text = "• Market entry requires localization strategy, navigating country-specific reimbursement systems, and overcoming trust deficit"
        p.font.size = Pt(14)
        
        p = tf.add_paragraph()
        p.text = "• Strategic acquisitions and partnerships critical for European market penetration"
        p.font.size = Pt(14)
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://drive.google.com/file/d/1zt3vNjOlugypsvPkywduyWNRPGTHrsoj/view?usp=drive_link"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_agenda():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Agenda"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        points = [
            "Market Overview & Key Metrics",
            "Regulatory Landscape Analysis", 
            "Competitive Positioning",
            "Access & Reimbursement Framework",
            "Strategic Opportunities & Risks",
            "Recommendations for Market Entry"
        ]
        
        for point in points:
            p = tf.add_paragraph()
            p.text = point
            p.font.size = Pt(16)
            p.level = 0
    
    def add_market_overview():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "European IVD Market: Substantial Opportunity with Complex Structure"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "Market Size & Structure"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Germany: €12.9B annual IVD expenditure, 2.6% of total healthcare costs"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Per capita spending: €150 annually in Germany"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• 55% growth in lab spending (2012-2022) vs 63% overall health expenditure growth"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Market dominated by large commercial laboratory networks"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "Key Market Trends"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Consolidation: Under 20% of clinics maintain own laboratory infrastructure"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• POCT growth driven by technical personnel shortages and infrastructure decline"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Therapeutic Drug Monitoring emerging as growth area"
        p.font.size = Pt(14)
        p.level = 0
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://journals.publisso.de/en/journals/gms/volume23/000337"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_regulatory_landscape():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Regulatory Landscape: IVDR Compliance and Geopolitical Barriers"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "IVDR Transition Challenges"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Transition to IVDR extended to 2027-2028 with stricter compliance requirements"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Certification process lengthened by 6-12 months, forcing out 30%+ of SMEs"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Environmental regulations (REACH) add compliance complexity for reagents"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "International Procurement Instrument (IPI)"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Effective June 2025: Bans Chinese companies from public contracts >€5M"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Limits Chinese component usage to 50% maximum in successful bids"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Exceptions only where no alternative suppliers exist"
        p.font.size = Pt(14)
        p.level = 0
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://drive.google.com/file/d/1zt3vNjOlugypsvPkywduyWNRPGTHrsoj/view?usp=drive_link [2] https://ec.europa.eu/commission/presscorner/detail/en/ip_25_1569"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_competitive_analysis():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Competitive Landscape: Chinese Suppliers Gaining Traction"
        
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9)
        height = Inches(4)
        
        table = slide.shapes.add_table(5, 4, left, top, width, height).table
        
        table.columns[0].width = Inches(2)
        table.columns[1].width = Inches(2)
        table.columns[2].width = Inches(2.5)
        table.columns[3].width = Inches(2.5)
        
        table.cell(0, 0).text = "Company"
        table.cell(0, 1).text = "Regulatory Progress"
        table.cell(0, 2).text = "Financial Strength"
        table.cell(0, 3).text = "Market Strategy"
        
        companies = [
            ["Mindray", "IVDR certified, HyTest acquisition", "11.58B CNY net profit (2023)", "Localization via acquisitions"],
            ["Snibe", "211 chemiluminescence IVDR certificates", "16.62% revenue growth (Q1 2024)", "Distributor network in 18 countries"],
            ["Autobio", "IVDR certified, Sekisui partnership", "Positive net profit growth", "Strategic partnerships"],
            ["Wondfo", "IVDR QMS certified, 150+ countries", "Global manufacturing footprint", "POCT focus, acquisitions"]
        ]
        
        for i, company in enumerate(companies, 1):
            for j, value in enumerate(company):
                table.cell(i, j).text = value
                table.cell(i, j).text_frame.paragraphs[0].font.size = Pt(10)
        
        for cell in table.iter_cells():
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(10)
                paragraph.alignment = PP_ALIGN.LEFT
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://en.caclp.com/industry-news/3530.html [2] https://en.caclp.com/industry-news/2785.html [3] https://en.caclp.com/industry-news/3349.html"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_registration_chart():
        fig, ax = plt.subplots(figsize=(8, 5))
        
        companies = ['Mindray', 'Maccura', 'Snibe', 'Other Chinese', 'Global Leaders']
        registrations = [3914, 2467, 1493, 2000, 19000]
        
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
        bars = ax.bar(companies, registrations, color=colors)
        
        ax.set_ylabel('Number of Registrations', fontsize=12)
        ax.set_title('Global Product Registrations: Chinese vs Global Leaders', fontsize=14)
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + 100,
                    f'{height:,}', ha='center', va='bottom', fontsize=10)
        
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', dpi=300)
        img_stream.seek(0)
        
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Product Registration Gap: Chinese Suppliers vs Global Leaders"
        
        left = Inches(1)
        top = Inches(1.5)
        slide.shapes.add_picture(img_stream, left, top, width=Inches(8))
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://en.caclp.com/industry-news/2365.html"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
        
        plt.close()
    
    def add_access_reimbursement():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Access & Reimbursement: Fragmented European Landscape"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "Country-Specific Systems"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Germany: EBM catalog for outpatient, DRG-based for inpatient"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• France: NABM system with innovative LAHN pathway"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Italy: LEA catalog with defined tariffs"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Spain: Global budget funding, no dedicated reimbursement"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• UK: NHS Payment Scheme, Medtech Funding Mandate"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "Key Challenges"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Fixed prices without inflation adjustment (Germany)"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Manufacturer-triggered reimbursement code creation"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• HTA requirements vary by country and region"
        p.font.size = Pt(14)
        p.level = 0
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://mtrconsult.com/general-market-access-landscape-vitro-diagnostic-tests-europe [2] https://journals.publisso.de/en/journals/gms/volume23/000337"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_budget_analysis():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Budget Analysis: Cost Pressures Drive Value Seeking"
        
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9)
        height = Inches(3)
        
        table = slide.shapes.add_table(5, 3, left, top, width, height).table
        
        table.columns[0].width = Inches(3)
        table.columns[1].width = Inches(3)
        table.columns[2].width = Inches(3)
        
        table.cell(0, 0).text = "Country"
        table.cell(0, 1).text = "Funding Mechanism"
        table.cell(0, 2).text = "Cost Pressure Factors"
        
        data = [
            ["Germany", "Global budgets, DRG-based", "Fixed prices, efficiency bonus"],
            ["France", "NABM with point values", "Innovation funding constraints"],
            ["Spain", "Regional global budgets", "No dedicated reimbursement"],
            ["UK", "NHS Payment Scheme", "Medtech Funding Mandate limits"]
        ]
        
        for i, row in enumerate(data, 1):
            for j, value in enumerate(row):
                table.cell(i, j).text = value
                table.cell(i, j).text_frame.paragraphs[0].font.size = Pt(10)
        
        for cell in table.iter_cells():
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(10)
                paragraph.alignment = PP_ALIGN.LEFT
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "Strategic Implications"
        p.font.size = Pt(16)
        p.font.bold = True
        p._p.getparent().remove(p._p)
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(1.5))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "• Laboratories operate under severe cost constraints, creating demand for cost-effective solutions"
        p.font.size = Pt(12)
        
        p = tf.add_paragraph()
        p.text = "• Chinese suppliers must demonstrate significant cost advantages over incumbents"
        p.font.size = Pt(12)
        
        p = tf.add_paragraph()
        p.text = "• Value proposition must align with hospital budget optimization needs"
        p.font.size = Pt(12)
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://mtrconsult.com/general-market-access-landscape-vitro-diagnostic-tests-europe [2] https://journals.publisso.de/en/journals/gms/volume23/000337"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_customer_receptivity():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Customer Receptivity: Trust Deficit vs Cost Advantage"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "Adoption Challenges"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Low brand recognition for high-end products in developed markets"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Market dominated by entrenched Western incumbents (Roche, Siemens, Abbott)"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Trust deficit hampers adoption of high-end IVD solutions"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "Competitive Advantages"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Strong price competitiveness and cost-effectiveness"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Performance-to-cost differentiation strategy"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Successful adoption in price-sensitive segments and emerging markets"
        p.font.size = Pt(14)
        p.level = 0
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://drive.google.com/file/d/1zt3vNjOlugypsvPkywduyWNRPGTHrsoj/view?usp=drive_link [2] https://en.caclp.com/industry-news/2365.html"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_strategic_opportunities():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Strategic Opportunities: Pathways to Market Success"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "Market Entry Strategies"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Utilize Eastern European countries (Hungary, Greece) as gateways to EU market"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Leverage German certification for enhanced brand credibility"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Focus on POCT and decentralized testing segments"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "Strategic Partnerships & Acquisitions"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Acquire European companies for local presence (Mindray-DiaSys model)"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Secure upstream raw materials through acquisitions (HyTest strategy)"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Establish local manufacturing to bypass IPI restrictions"
        p.font.size = Pt(14)
        p.level = 0
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://drive.google.com/file/d/1zt3vNjOlugypsvPkywduyWNRPGTHrsoj/view?usp=drive_link [2] https://hytest.fi/news/mindray-completes-acquisition-of-hytest"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_risks_challenges():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Risks & Challenges: Navigating Complex Barriers"
        
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9)
        height = Inches(4)
        
        table = slide.shapes.add_table(6, 2, left, top, width, height).table
        
        table.columns[0].width = Inches(4.5)
        table.columns[1].width = Inches(4.5)
        
        table.cell(0, 0).text = "Risk Category"
        table.cell(0, 1).text = "Specific Challenges"
        
        risks = [
            ["Regulatory", "IVDR compliance costs, IPI restrictions, REACH environmental rules"],
            ["Market Access", "Country-specific reimbursement systems, HTA requirements"],
            ["Competitive", "Entrenched incumbents, trust deficit in high-end segments"],
            ["Operational", "Supply chain localization requirements, skilled workforce shortages"],
            ["Geopolitical", "Trade tensions, reciprocity requirements, policy changes"]
        ]
        
        for i, risk in enumerate(risks, 1):
            for j, value in enumerate(risk):
                table.cell(i, j).text = value
                table.cell(i, j).text_frame.paragraphs[0].font.size = Pt(10)
        
        for cell in table.iter_cells():
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(10)
                paragraph.alignment = PP_ALIGN.LEFT
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://ec.europa.eu/commission/presscorner/detail/en/ip_25_1569 [2] https://mtrconsult.com/general-market-access-landscape-vitro-diagnostic-tests-europe"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_revenue_growth_chart():
        fig, ax = plt.subplots(figsize=(8, 5))
        
        companies = ['Mindray', 'Snibe', 'Dirui', 'Industry Average']
        growth_rates = [20, 16.62, 99.6, -53.26]
        
        colors = ['green' if x > 0 else 'red' for x in growth_rates]
        bars = ax.bar(companies, growth_rates, color=colors)
        
        ax.set_ylabel('Revenue Growth (%)', fontsize=12)
        ax.set_title('Chinese IVD Companies: Diverging Performance (2023-2024)', fontsize=14)
        ax.axhline(y=0, color='black', linestyle='-', alpha=0.3)
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + (5 if height > 0 else -8),
                    f'{height}%', ha='center', va='bottom' if height > 0 else 'top', 
                    fontsize=10, fontweight='bold')
        
        plt.tight_layout()
        
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', dpi=300)
        img_stream.seek(0)
        
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Financial Performance: Export-Focused Companies Outperforming"
        
        left = Inches(1)
        top = Inches(1.5)
        slide.shapes.add_picture(img_stream, left, top, width=Inches(8))
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://en.caclp.com/industry-news/2808.html [2] https://en.caclp.com/industry-news/2785.html"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
        
        plt.close()
    
    def add_recommendations():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Strategic Recommendations for Market Entry"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "Immediate Priorities (0-12 months)"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Establish local EU manufacturing/JVs to bypass IPI restrictions"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Secure full IVDR certification for complete product portfolios"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Develop country-specific HTA and reimbursement dossiers"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "Medium-Term Strategy (1-3 years)"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Acquire European distributors or establish direct sales force"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Build local service and support capabilities across key markets"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Develop innovative payment models for cost-constrained systems"
        p.font.size = Pt(14)
        p.level = 0
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://www.ainvest.com/news/frontier-medtech-navigating-china-eu-trade-tensions-investment-gains-2507/ [2] https://mtrconsult.com/general-market-access-landscape-vitro-diagnostic-tests-europe"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_conclusion():
        slide = prs.slides.add_slide(content_slide_layout)
        title = slide.shapes.title
        title.text = "Conclusion: Strategic Imperatives for European Success"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "Key Success Factors"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Localization is non-negotiable: manufacturing, service, and compliance"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Focus on cost-advantaged segments while building high-end credibility"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Navigate country-specific reimbursement systems with tailored strategies"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "Future Outlook"
        p.font.size = Pt(16)
        p.font.bold = True
        
        p = tf.add_paragraph()
        p.text = "• Chinese suppliers well-positioned for mid-market segments"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Regulatory compliance becoming key differentiator"
        p.font.size = Pt(14)
        p.level = 0
        
        p = tf.add_paragraph()
        p.text = "• Market share gains expected in cost-sensitive segments"
        p.font.size = Pt(14)
        p.level = 0
        
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(1))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = "[1] https://drive.google.com/file/d/1zt3vNjOlugypsvPkywduyWNRPGTHrsoj/view?usp=drive_link [2] https://www.ainvest.com/news/frontier-medtech-navigating-china-eu-trade-tensions-investment-gains-2507/"
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(100, 100, 100)
    
    add_title_slide()
    add_executive_summary()
    add_agenda()
    add_market_overview()
    add_regulatory_landscape()
    add_competitive_analysis()
    add_registration_chart()
    add_access_reimbursement()
    add_budget_analysis()
    add_customer_receptivity()
    add_strategic_opportunities()
    add_risks_challenges()
    add_revenue_growth_chart()
    add_recommendations()
    add_conclusion()
    
    prs.save('chinese_ivd_europe_expansion.pptx')

create_presentation()