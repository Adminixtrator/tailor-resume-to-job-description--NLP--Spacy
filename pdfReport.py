import pdfkit as pk


match_score = input("Enter match score: ")
ats_present = bool(input("ATS present?: "))
missing_skills_count = int(input("Missing skills: "))
example_missing_skill = input("Enter example input skill: ")
example_missing_skill_count = int(input("Enter example input skill count: "))
job_title = input("Enter Job title: ")
job_title_found = input("Job title found?: ")
education_match = bool(input("Education match?: "))
sections_headings_found = list(input("Enter found sessions: "))
dates_format_correct = bool("Date format correct?: ")
resume_length = int(input("Resume length: "))
suggested_word_count = int(input("Suggested wCount: "))
measureable_result = int(input("Measurable result: "))
job_level = input("Job level: ")
years_of_experience = int(input("Years of experience: "))
buzz_words_count = input("Buzz words: ")



path_wkhtmltopdf = r'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'
config = pk.configuration(wkhtmltopdf=path_wkhtmltopdf)
pk.from_file('pdf.html', 'python.pdf', configuration=config)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               