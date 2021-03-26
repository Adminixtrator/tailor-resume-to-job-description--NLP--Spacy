import fitz
import en_core_web_sm
from spacy.pipeline import EntityRuler
import jsonlines
from titlecase import titlecase
from docx import Document

nlp = en_core_web_sm.load()

def clean_text(corpus):
        cleaned_text = ""
        for i in corpus:
            cleaned_text = cleaned_text + i.lower().replace("'", "")
        return cleaned_text

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
        return content.strip(), content.strip().replace("\n", " ").replace(bullet_regex, "").replace("\t", " ")
    
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

    resume = './'+input("Resume Filename: ")
    vacature_raw_text, vacature_text = TextExtract.extract_text_from_txt('./'+input("JD Filename: "))
    missing_skills = []
    skills_log = []
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
    
    print(responseJson)
    