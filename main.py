import fitz
import en_core_web_sm
from spacy.pipeline import EntityRuler
import jsonlines
from titlecase import titlecase
from docx import Document
import pdfkit as pk
import re

nlp = en_core_web_sm.load()
webR = re.compile(r"(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'\".,<>?«»“”‘’]))")
comR = re.compile(r"(?i)company:\s*(?P<name>[A-Za-z\t .]+)")
dateR = re.compile(r'\w+\s\d{4}\s\w{2,4}\s\w+\s\w{0,4}|\w+\s\d{4}\s\D+\s\w+\s\w{0,4}')
dateFR = re.compile(r'\w+\s\d{4}\s\w{2,4}\s\w+\s\w{0,4}\s\(|\w+\s\d{4}\s\D+\s\w+\s\w{0,4}\s\(')
buzzWords = ["highly motivated", "strong leadership", "management skills", "extensive experience", "problem solver", "problem solving", "results-oriented", "hits the ground running", "process improvements", "ambitious", "seasoned", "customer driven", "extensive experience", "vast experience", "thought leader", "natural leader", "creative thinker", "savvy"]

# Education Degrees
EDUCATION = [
            'BE','B.E.', 'B.E', 'BS', 'B.S','C.A.','c.a.','B.Com','B. Com','M. Com', 'M.Com','M. Com .',
            'ME', 'M.E', 'M.E.', 'MS', 'M.S',
            'BTECH', 'B.TECH', 'M.TECH', 'MTECH',
            'PHD', 'phd', 'ph.d', 'Ph.D.','MBA','mba','graduate', 'post-graduate','5 year integrated masters','masters',
            'SSC', 'HSC', 'CBSE', 'ICSE', 'X', 'XII'
        ]


def clean_text(corpus):
        cleaned_text = ""
        for i in corpus:
            cleaned_text = cleaned_text + i.lower().replace("'", "")
        return cleaned_text

def checkEducation(text):
    try:
        for i in text.splitlines():
            if 'education' in i.lower().split(" ") and len(i.lower().split(" ")) == 1:
                return True
            elif 'certifications' in i.lower().split(" ") and len(i.lower().split(" ")) == 1:
                return True
            elif (any("education" in j for j in i.lower().split(" ")) and len(i.lower().split(" ")) == 1) or (any("certifications" in j for j in i.lower().split(" ")) and len(i.lower().split(" ")) == 1):
                return True
            elif any("education" in j for j in i.lower().split(" ")):
                return True
    except IndexError:
        return False, ''

def checkExperience(text):
    try:
        for i in text.splitlines(): 
            if 'experience' in i.lower().split(" ") and len(i.split(" ")) <= 2:
                return True
            elif 'projects' in i.lower().split(" ") and len(i.split(" ")) <= 2:
                return True
            elif (any("experience" in j for j in i.lower().split(" ")) and len(i.split(" ")) <= 2) or (any("projects" in j for j in i.lower().split(" ")) and len(i.split(" ")) <= 2):
                return True
    except IndexError:
        return False


class TextExtract():

    def extract_text_from_pdf(file):
        content = ""
        bullet_regex = '\xc2\xb7'
        cnt = 0
        doc = fitz.open(file)
        while cnt < doc.pageCount:
            page = doc.loadPage(cnt)
            content = content + page.getText("text")
            cnt+=1
        return content.strip(), clean_text(content.strip().replace("\n", " ").replace(bullet_regex, "").replace("\t", " "))

    def extract_text_from_docx(file):
        content = ""
        bullet_regex = '\xc2\xb7'
        doc = Document(file)
        for l in doc.paragraphs:
            content = content + '\n' + l.text
        return content.strip(), clean_text(content.strip().replace("\n", " ").replace(bullet_regex, "").replace("\t", " "))
    
    def extract_text_from_txt(file):
        with open(file) as job:
            text = job.read()  
        return text.strip(), clean_text(text.strip().replace("\n", " ").replace("\t", " "))
    

def add_newruler_to_pipeline(skill_pattern_path):
    new_ruler = EntityRuler(nlp).from_disk(skill_pattern_path)
    nlp.add_pipe(new_ruler, after='parser') 
    
def visualize_entity_ruler(entity_list, doc):
    options = {"ents": entity_list}
    displacy.render(doc, style='ent', options=options)


def create_skill_set(doc):
    return set([ent.label_[6:] for ent in doc.ents if 'skill' in ent.label_.lower()])

def create_skillset_dict(resume_texts):
    skillsets = [create_skill_set(resume_text) for resume_text in resume_texts]
    return skillsets


if __name__ == '__main__':

    # resume = './'+input("Resume Filename: ")
    resume = './input.pdf'
    # vacature_raw_text, vacature_text = TextExtract.extract_text_from_txt('./'+input("JD Filename: "))
    vacature_raw_text, vacature_text = TextExtract.extract_text_from_txt('./job_description.txt')
    missing_skills = []
    skills_log = []
    skills_logR = []
    responseJson = {}
    add_newruler_to_pipeline('./skill_patterns.jsonl')

    if resume.count(".pdf") > 0 and len(resume) <= 100:
        resume_raw_text, resume_text = TextExtract.extract_text_from_pdf(resume)
    elif resume.count(".doc") > 0 and len(resume) <= 100:
        resume_raw_text, resume_text = TextExtract.extract_text_from_docx(resume)
    elif resume.count(".txt") > 0 and len(resume) <= 100:
        resume_raw_text, resume_text = TextExtract.extract_text_from_txt(resume)

    skillset_dict = create_skillset_dict([nlp(resume_text)])
    vacature_skillset = create_skill_set(nlp(vacature_text))
    responseJson['jobDesc'] = vacature_raw_text
    responseJson['resumeText'] = resume_raw_text

    if len(vacature_skillset) < 1:
        responseJson['skillsInJD'] = False
        responseJson['keywords'] = []
        responseJson['missingKeywords'] = []
    else:
        responseJson['skillsInJD'] = True
        pct_match = round(len(vacature_skillset.intersection(skillset_dict[0])) / len(vacature_skillset) * 100, 2)
        responseJson['percentageMatch'] = pct_match
        for skill in vacature_skillset:
            # ------------- PDF REPORT ASSET -----------------------
            skills_logR.append({'skill': skill.replace("-", " "), 'count': resume_text.lower().count(skill.replace("-", " "))})
            if vacature_text.lower().count(skill.replace("-", " ")) == 0:
                if len(skill.replace("-", " ")) >= 5 and len(skill.split("-")) > 1:
                    skills_log.append({'skill': skill.replace("-", " "), 'keySkill': True, 'count': 1})
                else:
                    skills_log.append({'skill': skill.replace("-", " "), 'keySkill': False, 'count': 1})
            else:
                if vacature_text.lower().count(skill.replace("-", " ")) == 1 and len(skill.split("-")) > 1 and len(skill.replace("-", " ")) >= 7:
                    skills_log.append({'skill': skill.replace("-", " "), 'keySkill': True, 'count': vacature_text.lower().count(skill.replace("-", " "))})
                elif vacature_text.lower().count(skill.replace("-", " ")) > 1:
                    skills_log.append({'skill': skill.replace("-", " "), 'keySkill': True, 'count': vacature_text.lower().count(skill.replace("-", " "))})
                else:
                    skills_log.append({'skill': skill.replace("-", " "), 'keySkill': False, 'count': vacature_text.lower().count(skill.replace("-", " "))})

            if skill not in vacature_skillset.intersection(skillset_dict[0]):
                if vacature_text.lower().count(skill.replace("-", " ")) == 0:
                    if len(skill.replace("-", " ")) >= 5 and len(skill.split("-")) > 1:
                        missing_skills.append({'skill': skill.replace("-", " "), 'keySkill': True, 'count': 1})
                    else:
                        missing_skills.append({'skill': skill.replace("-", " "), 'keySkill': False, 'count': 1})
                else:
                    if vacature_text.lower().count(skill.replace("-", " ")) == 1 and len(skill.split("-")) > 1 and len(skill.replace("-", " ")) >= 7:
                        missing_skills.append({'skill': skill.replace("-", " "), 'keySkill': True, 'count': vacature_text.lower().count(skill.replace("-", " "))})
                    elif vacature_text.lower().count(skill.replace("-", " ")) > 1:
                        missing_skills.append({'skill': skill.replace("-", " "), 'keySkill': True, 'count': vacature_text.lower().count(skill.replace("-", " "))})
                    else:
                        missing_skills.append({'skill': skill.replace("-", " "), 'keySkill': False, 'count': vacature_text.lower().count(skill.replace("-", " "))})
        
        responseJson['keywords'] = skills_log
        responseJson['pdfReportRKeywords'] = skills_logR
        responseJson['missingKeywords'] = missing_skills
    
    nlp = en_core_web_sm.load()
    add_newruler_to_pipeline('./job_titles.jsonl')
    job_tile_resume_dict = create_skillset_dict([nlp(resume_text)])
    job_tile_JD_dict = create_skill_set(nlp(vacature_text))

    if len(job_tile_JD_dict) < 1:
        responseJson['jobTitleInJD'] = False
    else:
        responseJson['jobTitleInJD'] = True
        responseJson['jobTitle'] = titlecase(str(max(list(job_tile_JD_dict), key=len)).replace("-", " "))
        if list(job_tile_JD_dict.intersection(job_tile_resume_dict[0])).count(str(max(list(job_tile_JD_dict), key=len))) > 0:
            responseJson['jobTitleInR'] = True
        else:
            responseJson['jobTitleInR'] = False
    
    # print(responseJson)


    # --------------------- PDF REPORT ---------------------------
    if responseJson['skillsInJD'] == True:
        pdf_report_html = """ 
        <!DOCTYPE html>
        <head>

        </head>
        <body style="font-family: 'Arial';">
        <br /><hr /><br />
        <h1 style="font-size: 3rem; text-align: center; font-weight: bolder;">Match Rate: """+str(responseJson['percentageMatch'])+""" %
        </h1>
        <h2 style="font-size: 1.5rem; text-align: center;">RESUME SCAN REPORT</h2><br /><br />

        <div style="padding-left: 5%; padding-right: 5%;">
        <h4>ATS FINDINGS</h4><br />
        <div style="margin-left: 2%;">

        <h3>ATS TIPS</h3>
        <p>"""
        jobUrlFound = False;
        companyNameFound = False;
        companyName = " Add Company Name"
        jobUrl = " Add web address for this job"
        try:
            jobUrl = list(webR.findall(responseJson['jobDesc'])[0])[0]
            jobUrlFound = True;
        except:
            pass

        try:
            companyName = comR.findall(responseJson['jobDesc'])[0]
            companyNameFound = True;
        except:
            pass
        
        if jobUrlFound == True or companyNameFound == True:
            pdf_report_html = pdf_report_html + '<img src="check.png" height="20" style="margin:  0 10px 0 10px;">'
        else:
            pdf_report_html = pdf_report_html + '<img src="cross.png" height="20" style="margin:  0 10px  0 10px;">'
        
        pdf_report_html = pdf_report_html + """ Adding this job's company name and(or) web address can help us provide you ATS- specific tips. 
        You can add the company name with the flag <i><b>Company:</b> JobsGig</i>, and add URL by pasting it at the top or bottom of the job description.</p>
        <p><b>Company:</b> """ 
        
        if companyNameFound == True:
            pdf_report_html = pdf_report_html + companyName + " | <b>URL:</b>"
        elif companyNameFound == False:
            pdf_report_html = pdf_report_html + "Add Company Name | <b>URL:</b>"

        if jobUrlFound == True:
           pdf_report_html = pdf_report_html + "<a href=\"" + jobUrl + "\" style=\"text-decoration\": none;\"> View site</a></p>"
        elif jobUrlFound == False:
           pdf_report_html = pdf_report_html + jobUrl + "</p>"
        pdf_report_html = pdf_report_html + """
        <br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />

        <h3>SKILLS AND KEYWORDS</h3>
        <p>
        """
        high_value_skills = []
        for i in responseJson['missingKeywords']:
            if i["count"] >= 2:
                high_value_skills.append(i)

        if len(high_value_skills) > 0:
            pdf_report_html = pdf_report_html + """<img src="cross.png" height="20" style="margin:  0 10px  0 10px;"> """
        else:
            pdf_report_html = pdf_report_html + """<img src="check.png" height="20" style="margin:  0 10px  0 10px;"> """

        if len(high_value_skills) > 0:
            pdf_report_html = pdf_report_html + """
            You are missing <b>"""+str(len(high_value_skills))+""" important high-value skills</b> on your resume.<p style="font-size: 0.75rem;">
            <p>For example, """+str(high_value_skills[0]["skill"])+""" appears on the job description """+str(high_value_skills[0]["count"])+""" times and is not on your resume. 
            <b>You are additionally missing some other hard and soft skills.</b></p>
            <br />Review your <b>missing skills below.</b></p>
            <br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />

            <h3>JOB TITLE MATCH</h3>
            <p>
            """
        else:
            pdf_report_html = pdf_report_html + """
            You are not missing any important high-value skills on your resume.<p style="font-size: 0.75rem;">
            <br />Review your <b>missing skills below.</b></p>
            <br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />

            <h3>JOB TITLE MATCH</h3>
            <p>
            """
        if responseJson['jobTitleInJD'] == True and responseJson['jobTitleInR'] == True:
            pdf_report_html = pdf_report_html + """
            <img src="check.png" height="20" style="margin:  0 10px  0 10px;"> The <i>'"""+str(responseJson['jobTitle'])+"""
            '</i> job title provided or found in the job description was found in your resume. 
            We recommend having the exact title of the job for which you're applying in your resume. This ensures you'll be 
            found when a recruiter searches by job title. If you haven't held this position before, include it as part of 
            your summary statement.<br />Incorrect job title in the job description?</p>
            <br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />

            <h3>EDUCATION MATCH</h3>
            <p>
            """
        elif responseJson['jobTitleInJD'] == True and responseJson['jobTitleInR'] == False:
            pdf_report_html = pdf_report_html + """
            <img src="cross.png" height="20" style="margin:  0 10px  0 10px;"> The <i>'"""+str(responseJson['jobTitle'])+"""
            '</i> job title provided or found in the job description was not found in your resume. 
            We recommend having the exact title of the job for which you're applying in your resume. This ensures you'll be 
            found when a recruiter searches by job title. If you haven't held this position before, include it as part of 
            your summary statement.<br />Incorrect job title in the job description?</p>
            <br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />

            <h3>EDUCATION MATCH</h3>
            <p>
            """
        elif responseJson['jobTitleInJD'] == False:
            pdf_report_html = pdf_report_html + """
            <img src="check.png" height="20" style="margin:  0 10px  0 10px;"> The job title was not provided or found in the job description. 
            We recommend having the exact title of the job for which you're applying in your resume. This ensures you'll be 
            found when a recruiter searches by job title. If you haven't held this position before, include it as part of 
            your summary statement.<br />Incorrect job title in the job description?</p>
            <br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />

            <h3>EDUCATION MATCH</h3>
            <p>
            """
            
        pdf_report_html = pdf_report_html + """<img src="check.png" height="20" style="margin:  0 10px  0 10px;"> This job doesn't specify a preferred degree.</p>
        <br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br /><br />
        
        <h3>SECTION HEADINGS</h3>
        <p>"""

        if checkExperience(responseJson['resumeText']) == True:
            pdf_report_html = pdf_report_html + '<p><img src="check.png" height="20" style="margin:  0 10px  0 10px;"> We found the work experience section in your resume.</p>'
        else: 
            pdf_report_html = pdf_report_html + '<p><img src="cross.png" height="20" style="margin:  0 10px  0 10px;"> We did not find the work experience section in your resume.</p>'

        if checkEducation(responseJson['resumeText']) == True:
            pdf_report_html = pdf_report_html + '<p><img src="check.png" height="20" style="margin:  0 10px  0 10px;"> We found the education section in your resume.</p>'
        else: 
            pdf_report_html = pdf_report_html + '<p><img src="cross.png" height="20" style="margin:  0 10px  0 10px;"> We did not find the education section in your resume.</p>'
        
        pdf_report_html = pdf_report_html + """<p>If you've included the section(s) and the tool hasn't picked it up, please make sure you've named your section titles correctly. 
        Use conventional section titles like <b>'Experience'</b>, <b>'Projects'</b> or <b>'Work'</b>.</p>
        <br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />

        <h3>DATE FORMATTING</h3>
        <p>
        """

        dates = []
        for i in dateR.findall(resume_text):
            dates.append(i.split(" ")[1])
        dtFormat = [True if len(dateFR.findall(resume_text)) > 0 else False][0]                        

        if dtFormat == True:
            pdf_report_html = pdf_report_html+ '<img src="check.png" height="20" style="margin:  0 10px  0 10px;"> The dates in your work experience section are properly formated.'
        else: 
            pdf_report_html = pdf_report_html+ '<img src="cross.png" height="20" style="margin:  0 10px  0 10px;"> The dates in your work experience section are not properly formated.'

        pdf_report_html = pdf_report_html + """</p><br /><br />

        </div><br /><br />

        <h4>RECRUITER FINDINGS</h4>
        <div style="margin-left: 2%;">

        <h3>WORD COUNT</h3>
        <p>"""

        if 6*len(responseJson['resumeText'].splitlines()) < 1000:
            pdf_report_html = pdf_report_html + '<img src="check.png" height="20" style="margin:  0 10px  0 10px;"> There are ' + str(6*len(responseJson['resumeText'].splitlines())) + ' words in your resume, which is under the suggested 1000 word count for relevance and ease of reading reasons.</p>'
        if 6*len(responseJson['resumeText'].splitlines()) > 1000:
            pdf_report_html = pdf_report_html + '<img src="cross.png" height="20" style="margin:  0 10px  0 10px;"> There are ' + str(6*len(responseJson['resumeText'].splitlines())) + ' words in your resume, which is above the suggested 1000 word count for relevance and ease of reading reasons.</p>'

        pdf_report_html = pdf_report_html + """<br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />

        <h3>MEASURABLE RESULTS</h3>
        <p><img src="warning.png" height="20" style="margin:  0 10px  0 10px;"> We only found only a few mentions of measurable results in your resume. Consider adding at least 5 specific achievements or impact you had in your job 
        (e.g. time saved, increase in sales, etc).</p><br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />

        <h3>JOB LEVEL MATCH</h3>
        <p><img src="check.png" height="20" style="margin:  0 10px  0 10px;"> You are applying to a(n) junior level role. Given your 3 years of experience, this role is a great fit for your experience.</p>
        <br /><hr style="background-color: #eeefff; height: 1px; border: none;"><br />
        
        <h3>WORDS TO AVOID</h3>
        <p>
        """

        bzW = []
        for i in buzzWords:
            if resume_text.lower().count(i) > 0:
                bzW.append(i)

        if len(bzW) == 0:
            pdf_report_html = pdf_report_html + '<img src="check.png" height="20" style="margin:  0 10px  0 10px;"> We did not find vague buzzwords on your resume.</p>'
        else: 
            pdf_report_html = pdf_report_html + '<img src="warning.png" height="20" style="margin:  0 10px  0 10px;"> We found ' + str(len(bzW)) + ' vague buzzwords on your resume (' + ", ".join(bzW) + ''').</p>
            <p>Phrases like these are a red flag to employers and should be removed. They add fluff to your resume and are often too vague to communicate any actual value.</p>'''

        pdf_report_html = pdf_report_html + '''<br /><br />

        </div><br /><br />

        <h1 style="font-size: 1.3rem;">SKILLS [ HIGH IMPORTANCE ]</h1>
        <p>Skills are typically learned through education or work experience, such as proficiency with specific software, tools, or 
        specialized processes. Try to match your resume skills exactly to those in the job description. Skills marked with <img src="cross.png" height="10" style="margin: 0 10px 0 10px;"> were found in 
        the job description but not in your resume. 
        Prioritize skills that appear most frequently in the job description.</p>
        <hr style="background-color: skyblue; height: 5px; border: none;"><br />
        
        <h4 style="color: skyblue; text-align: center;">Skills Comparision</h4><br />

        <div style="width: 100%; margin: 0 auto;">
            <h4 style="display: inline-block; width: 25%; text-align: end; margin: 0 auto;">SKILL</h4>
            <h4 style="display: inline-block; width: 25%; text-align: end; margin: 0 auto;">RESUME</h4>
            <h4 style="display: inline-block; width: 32%; text-align: end; margin: 0 auto;">JOB DESCRIPTION</h4>
        </div>
        <hr style="background-color: #a9a9a9; height: 5px; border: none;">'''

        for key in responseJson['keywords']:
            pdf_report_html = pdf_report_html + '''
            <hr style="background-color: #eeefff; height: 1px; border: none;">
            <div style="width: 70%; margin-left: 13%;">
            <p style="display: inline-block; width: 25%; text-align: center;">''' + key["skill"] + '''</p>
            <p style="display: inline-block; width: 40%; text-align: center;">'''
            
            if key not in responseJson['missingKeywords']:
                pdf_report_html = pdf_report_html + str(key["count"])
            else:
                pdf_report_html = pdf_report_html + '<img src="cross.png" height="10" style="margin: 0 10px 0 10px;">'
            pdf_report_html = pdf_report_html + '''</p>
            <p style="display: inline-block; width: 32%; text-align: center;">''' + str(key["count"]) + '''</p>
            </div>'''

        

        f = open("index.html", mode="w")
        f.write(pdf_report_html)
        f.close()

        path_wkhtmltopdf = r'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'
        config = pk.configuration(wkhtmltopdf=path_wkhtmltopdf)
        pk.from_file('index.html', 'index.pdf', configuration=config)                             
    
