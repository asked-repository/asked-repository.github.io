"""Build the repository's data tables (surveys, concepts, variables) as CSV.

Run once to create the initial tables. After that, the CSVs are the master
files: edit them directly (Excel -> Save as CSV UTF-8).
"""
import csv, re, sys, os
import openpyxl

OUT = sys.argv[1] if len(sys.argv) > 1 else "site/data"
XLSX = sys.argv[2] if len(sys.argv) > 2 else "Codebook_UC3M (40dB) 05022026.xlsx"
os.makedirs(OUT, exist_ok=True)

# ---------------------------------------------------------------- themes
THEMES = [
    ("sociodem", "Socio-demographics"),
    ("ses", "Socio-economic status"),
    ("electoral", "Electoral behaviour"),
    ("polatt", "Political attitudes & engagement"),
    ("trust", "Trust"),
    ("eu", "Attitudes towards the EU"),
    ("gender", "Attitudes towards gender equality"),
    ("sexgender", "Gender identity, sexuality & morality policy"),
    ("media", "Media & social media"),
    ("ai", "AI & digital technologies"),
    ("personal", "Personal life, values & socialisation"),
    ("design", "Design features (experiments, weights)"),
]

# id | label | themes | search synonyms (EN / ES / CA)
CONCEPTS = [
    ("gender", "Gender / sex", "sociodem", "gender; sex; man; woman; men; women; género; sexo; hombre; mujer; gènere"),
    ("age", "Age", "sociodem", "age; cohort; generation; young; youth; edad; jóvenes; edat"),
    ("marital", "Marital & relationship status", "sociodem;personal", "marital status; married; partner; couple; relationship; dating; single; estado civil; pareja"),
    ("children", "Children & family", "sociodem", "children; kids; parenthood; parents; fertility; family; daughters; hijos; hijas; fills"),
    ("siblings", "Siblings", "sociodem", "siblings; brothers; sisters; hermanos; germans"),
    ("sexual_orientation", "Sexual orientation", "sociodem;sexgender", "sexual orientation; LGBT; gay; lesbian; bisexual; heterosexual; orientación sexual"),
    ("ethnicity", "Ethnicity & migration background", "sociodem", "ethnicity; race; ethnic; migrant; immigrant; migration background; citizenship; country of birth; nationality; origen; inmigrante; nacionalidad"),
    ("residence", "Place of residence (country, region, urban/rural)", "sociodem", "residence; region; rural; urban; city; territory; country; postcode; province; comunidad autónoma; provincia"),
    ("religion", "Religion & religiosity", "sociodem;personal", "religion; religiosity; religious; church attendance; catholic; religión; religiosidad"),
    ("education", "Education", "ses", "education; educational attainment; schooling; degree; university; educación; estudios; nivel educativo"),
    ("occupation", "Occupation / social class", "ses", "occupation; class; social class; ISCO; job; profession; ocupación; clase social"),
    ("employment", "Employment status", "ses", "employment; unemployment; unemployed; labour market status; work status; retired; student; empleo; paro; situación laboral"),
    ("income", "Income", "ses", "income; household income; earnings; wealth; ingresos; renta; salario"),
    ("social_origin", "Social origin (parents' occupation)", "ses", "social origin; parents occupation; father; mother; social mobility; class origin; origen social"),
    ("subjective_economy", "Personal economic situation", "ses", "personal economic situation; household finances; egotropic; economic hardship; situación económica personal"),
    ("vote_choice", "Vote choice / vote intention", "electoral", "vote choice; vote intention; party choice; recalled vote; voting behaviour; party vote; voto; intención de voto; recuerdo de voto"),
    ("turnout", "Turnout / abstention", "electoral", "turnout; abstention; voted; electoral participation; participación electoral; abstención"),
    ("ptv", "Party evaluations & propensity to vote", "electoral", "party like dislike; party thermometer; propensity to vote; PTV; party sympathy; affective polarization; polarisation; valoración de partidos"),
    ("leader_eval", "Leader / candidate evaluations", "electoral", "leader; leaders; candidate; candidates; leader evaluation; likeability; valoración de líderes; candidatos"),
    ("campaign", "Campaign exposure & mobilisation", "electoral", "campaign; electoral campaign; debates; party contact; canvassing; mobilisation; mobilization; campaña"),
    ("gov_eval", "Government performance", "electoral", "government approval; government performance; retrospective; evaluación del gobierno"),
    ("economic_perceptions", "National economic perceptions", "electoral", "economy; sociotropic; economic evaluation; inflation; prices; unemployment perceptions; economía; precios"),
    ("left_right", "Ideology (left–right)", "polatt", "ideology; left right; left-right; self-placement; ideología; izquierda derecha"),
    ("political_interest", "Political interest", "polatt", "political interest; interest in politics; interés por la política"),
    ("efficacy", "Political efficacy", "polatt", "efficacy; external efficacy; political influence; voice; eficacia política"),
    ("democracy_satisfaction", "Satisfaction with democracy", "polatt", "satisfaction with democracy; democratic satisfaction; satisfacción con la democracia"),
    ("regime_preferences", "Regime preferences / support for democracy", "polatt", "regime preferences; democracy support; strong leader; technocracy; experts; direct democracy; authoritarianism; military rule; stealth democracy; populism"),
    ("corruption", "Perceived corruption", "polatt", "corruption; corrupción"),
    ("taxation", "Taxation & redistribution", "polatt", "tax; taxes; taxation; redistribution; progressive taxation; tax evasion; impuestos; redistribución"),
    ("deservingness", "Deservingness / meritocracy", "polatt", "deservingness; meritocracy; deserve; fairness; merecimiento"),
    ("climate", "Climate change & energy", "polatt", "climate change; environment; energy; renewables; cambio climático; energía; medio ambiente"),
    ("incivility", "Political incivility, conflict & compromise", "polatt", "incivility; uncivil; political conflict; compromise; negative campaigning; polarization; crispación"),
    ("political_violence", "Political violence", "polatt", "political violence; intimidation; violencia política"),
    ("protest", "Protest & non-electoral participation", "polatt", "protest; demonstration; political participation; mobilisation; 8M; manifestación"),
    ("trust_political", "Trust in national political institutions & actors", "trust", "political trust; institutional trust; trust in parliament; trust in government; trust politicians; distrust; confianza política; confianza en las instituciones"),
    ("trust_eu", "Trust in the EU", "trust;eu", "EU trust; trust in the EU; EU distrust; trust European Commission; trust European Parliament; confianza en la UE"),
    ("social_trust", "Social / interpersonal trust", "trust", "social trust; interpersonal trust; generalised trust; generalized trust; trust in people; confianza social"),
    ("eu_integration", "EU integration & EU competences", "eu", "European integration; EU integration; EU competences; EU policy; federalism; Europeanisation; integración europea"),
    ("eu_reform", "EU democratic reform", "eu", "EU reform; democratic deficit; Spitzenkandidat; transnational lists; EU referendum"),
    ("eu_knowledge", "EU / political knowledge", "eu;polatt", "knowledge; political knowledge; EU knowledge; informed; conocimiento político"),
    ("identity_attachment", "National & European identity / attachment", "eu", "identity; attachment; national identity; European identity; nationalism; identidad; nacionalismo"),
    ("zero_sum", "Zero-sum perceptions", "eu", "zero-sum; zero sum; relative gains; winners and losers"),
    ("gender_attitudes", "Gender attitudes & sexism", "gender", "gender attitudes; sexism; hostile sexism; benevolent sexism; modern sexism; gender roles; gender equality; machismo; actitudes de género; igualdad de género"),
    ("gender_backlash", "Anti-feminism / gender backlash", "gender", "backlash; anti-feminism; antifeminism; gone too far; men disadvantaged; antifeminismo"),
    ("feminism", "Feminism & the feminist movement", "gender", "feminism; feminist; feminist movement; women's movement; 8M; feminismo"),
    ("gender_policy", "Gender equality policies", "gender", "gender policy; gender quotas; quotas; parity; parental leave; equal pay; cuotas; paridad; permisos"),
    ("gbv", "Gender-based violence", "gender", "gender-based violence; GBV; violence against women; Istanbul Convention; sexual consent; violencia de género"),
    ("women_representation", "Women in politics / candidate gender", "gender", "women in politics; women candidates; female politicians; political representation; candidate gender; mujeres en política"),
    ("discrimination", "Perceived discrimination", "gender", "discrimination; unfair treatment; discriminación"),
    ("traditional_values", "Traditional & authoritarian values", "personal;gender", "traditional values; family values; authoritarianism; child qualities; valores tradicionales"),
    ("gender_identity", "Gender identity (masculinity / femininity)", "sexgender", "masculinity; femininity; gender identity; gender salience; masculinidad; feminidad"),
    ("lgbtiq", "LGBTIQ+ & trans rights", "sexgender", "LGBT; LGBTIQ; trans; transgender; gender self-determination; same-sex adoption; trans law; ley trans"),
    ("abortion", "Abortion", "sexgender", "abortion; reproductive rights; aborto"),
    ("morality_policy", "Morality policy (surrogacy, sex work, euthanasia)", "sexgender", "morality policy; surrogacy; sex work; prostitution; euthanasia; gestación subrogada; trabajo sexual"),
    ("news_media", "News & media use", "media", "news; media use; information sources; newspapers; television; radio; consumo de medios; noticias"),
    ("social_media", "Social media & digital habits", "media", "social media; internet use; online; digital use; screen time; redes sociales; TikTok; Instagram; Twitter"),
    ("influencers", "Influencers & content creators", "media", "influencers; streamers; content creators; YouTubers; parasocial; creadores de contenido"),
    ("ai_attitudes", "Attitudes towards AI", "ai", "artificial intelligence; AI; AI attitudes; AI emotions; AI risks; inteligencia artificial; IA"),
    ("ai_use", "Use of AI tools", "ai", "AI use; ChatGPT; generative AI; Copilot; Gemini; uso de IA"),
    ("ai_governance", "AI in government & AI regulation", "ai", "AI regulation; algorithmic governance; AI in public administration; automated decision-making; regulación IA"),
    ("digital_rights", "Digital technologies & digital rights", "ai", "digital rights; privacy; data protection; digitalisation; digitalization; technology trust; facial recognition; cybersecurity"),
    ("life_satisfaction", "Life satisfaction & wellbeing", "personal", "life satisfaction; wellbeing; happiness; satisfacción vital"),
    ("personality", "Personality traits", "personal", "personality; big five; traits; risk taking; empathy; leadership; personalidad"),
    ("socialisation", "Socialisation (school, adolescence, friendships)", "personal", "socialisation; socialization; adolescence; school; friendships; friends; peers; single-sex schooling; socialización; amistades"),
    ("leisure", "Leisure & cultural consumption", "personal", "leisure; culture; sport; reading; cultural consumption; ocio"),
    ("care", "Care work", "personal;gender", "care; care work; caregiving; childcare; cuidados"),
    ("conjoint", "Conjoint experiment", "design", "conjoint; conjoint experiment; experiment"),
    ("vignette", "Vignette / split-ballot experiment", "design", "vignette; survey experiment; split ballot; split-sample; experiment"),
    ("weights", "Survey weights", "design", "weights; weight; weighting"),
]

# ---------------------------------------------------------------- surveys
TBC = "To confirm"
SURVEYS = [
    dict(id="express2",
         title="EXPRESS2 cross-national survey on trust in the EU",
         project="EXPRESS2 (Horizon Europe) — Universidad Carlos III de Madrid",
         institution="Universidad Carlos III de Madrid",
         countries="Austria;Belgium;Bulgaria;Croatia;Czech Republic;Denmark;Estonia;Finland;France;Germany;Greece;Hungary;Ireland;Italy;Latvia;Netherlands;Poland;Portugal;Romania;Serbia;Slovakia;Spain;Sweden;Turkey;United Kingdom",
         field_dates=TBC, year="", sample_size=TBC, population="Adults (18+) residing in each country",
         mode="Online panel (40dB)", languages="11 languages incl. English, German, Spanish, Italian, Portuguese",
         weights="Yes — gender, age and region",
         topics="Trust and distrust in EU institutions; EU knowledge and competences; EU democratic reform; gender equality and gender-based violence; AI and digital rights",
         experiments="Split-ballot on trust scales (0–10 vs −5/+5); conjoint on who designs EU policy (incl. AI)",
         status="embargoed", embargo_until="", data_link="", open_to_collaboration=TBC,
         citation=TBC, contact_name=TBC, contact_email="", questionnaire_file="", notes="Codebook dated 5 February 2026."),
    dict(id="aidem",
         title="Democracy in the Age of Artificial Intelligence",
         project="Democracy in the Age of Artificial Intelligence",
         institution=TBC,
         countries="Spain",
         field_dates=TBC, year="", sample_size=TBC, population=TBC,
         mode="Online panel", languages="Spanish", weights=TBC,
         topics="Generative AI use and attitudes; AI in public decision-making; influencers and social media; gender attitudes and feminism; masculinity/femininity; socialisation; taxation of the rich; morality policy",
         experiments="Conjoint on AI involvement in policy processes; vignette on taxing windfall incomes",
         status="embargoed", embargo_until="", data_link="", open_to_collaboration=TBC,
         citation=TBC, contact_name="Alejandro Tirado Castro", contact_email="", questionnaire_file="", notes=""),
    dict(id="twiceasgood",
         title="Twice as Hard, Half as Good? Women Candidates on the Campaign Trail — Post-electoral Study in Spain 2023",
         project="Twice as Hard, Half as Good? (University of Exeter)",
         institution="University of Exeter",
         countries="Spain",
         field_dates="After the 23 July 2023 general election (exact dates to confirm)", year="2023", sample_size=TBC,
         population="Adults (18+) residing in Spain",
         mode="Online panel (Netquest)", languages="Spanish", weights=TBC,
         topics="2023 general election campaign and vote; leader evaluations and debate tone; gender identity and sexism; gender policies; feminism and 8M; incivility and political violence; climate change",
         experiments="",
         status="embargoed", embargo_until="", data_link="", open_to_collaboration=TBC,
         citation=TBC, contact_name=TBC, contact_email="", questionnaire_file="", notes=""),
]

# ------------------------------------------------------------- variables
# (survey, var, question_en, question_original, response_scale, tags)
V = []
def add(s, var, en, orig, scale, tags):
    V.append(dict(survey_id=s, variable=var, question_en=en, question_original=orig,
                  response_scale=scale, concepts=tags))

# ---- EXPRESS2: grouped from the Excel codebook ----
wb = openpyxl.load_workbook(XLSX, data_only=True)
rows = [(str(r[0]).strip(), str(r[1] or "").strip(), str(r[2] or "").strip())
        for r in wb.active.iter_rows(min_row=2, values_only=True) if r[0]]
lab = {v: l for v, l, _ in rows}
vals = {v: x for v, _, x in rows}

def items(pattern):
    out = []
    for v, l, _ in rows:
        if re.fullmatch(pattern, v):
            out.append((v, l.split(" - ")[0].strip() if " - " in l else l))
    return out

def stem(v):
    l = lab[v]
    return l.split(" - ", 1)[1].strip() if " - " in l else l

def scale(v, n=6):
    parts = [p.strip() for p in vals.get(v, "").split("\n") if p.strip()]
    parts = [re.sub(r'^-?\d+\s*=\s*"?|"$', "", p) for p in parts]
    if not parts: return ""
    if len(parts) > n: return " … ".join([parts[0], parts[-1]]) + f" ({len(parts)} categories)"
    return " / ".join(parts)

def battery(var_label, pattern, en, tags, n=6, extra=""):
    its = items(pattern)
    first = its[0][0]
    orig = stem(first) + " Items: " + "; ".join(i for _, i in its) + (extra and " " + extra)
    add("express2", var_label, en, orig, scale(first, n), tags)

def single(v, en, tags, n=6, var_label=None):
    add("express2", var_label or v, en, lab[v], scale(v, n), tags)

single("q0", "Country of residence", "residence", n=0)
add("express2", "region_*", "Region of residence (NUTS-type regions, one variable per country)", "In which region do you reside?", "Country-specific list", "residence")
single("q1", "Age (years); also recoded in age_r", "age", var_label="q1, age_r")
single("q2", "Gender identification", "gender")
single("q3", "Highest level of education completed (ISCED)", "education", n=0)
single("q4", "Employment status", "employment", n=0)
single("q5", "Occupation (ISCO major groups)", "occupation", n=0)
add("express2", "q6, q6_*", "Household income after taxes (country-specific bands, harmonised into deciles)", lab["q6_sp"], "Deciles (1st–10th) + country-specific bands", "income")
single("q7", "Religious affiliation", "religion", n=0)
single("q8", "Frequency of attending religious services", "religion", n=0)
single("q9", "Marital status", "marital", n=0)
battery("q10_1–q10_14", r"q10_\d+", "Ethnic self-identification (multiple choice)", "ethnicity", n=0)
single("q11", "Migration background and citizenship", "ethnicity", n=0)
single("q12", "Type of area of residence (rural to large city)", "residence")
add("express2", "q13_*, q13", "Vote intention in European Parliament elections (country-specific party lists, recoded into party families)", lab["q13_sp"], "Country-specific party list; party family recode", "vote_choice")
add("express2", "q14_*, q14", "Vote intention in national legislative elections (country-specific party lists, recoded into party families)", lab["q14_sp"], "Country-specific party list; party family recode", "vote_choice")
single("q15", "Left–right self-placement (0–10)", "left_right", n=0)
single("q16", "Life satisfaction (0–10)", "life_satisfaction", n=0)
add("express2", "q17", "Daily time spent following news about the EU", lab["q17"], "Time (open)", "news_media;eu_integration")
battery("q18_1–q18_7", r"q18_\d+", "Sources of political information (newspapers, TV, radio, podcasts, websites, social media, friends/family)", "news_media;social_media", n=0)
battery("q19_1–q19_2", r"q19_\d", "External efficacy: influence on politics in own country and in the EU", "efficacy;eu_integration")
single("q23", "Self-assessed knowledge of the EU and its institutions", "eu_knowledge", n=0)
add("express2", "q24_1–q24_7", "EU knowledge quiz (7 items: member states, Commission president, Schengen, EP elections, appointment of Commissioners, election of MEPs)",
    "; ".join(l for _, l in items(r"q24_\d")), "Multiple choice with one correct answer + Don't know", "eu_knowledge")
battery("q25_1–q25_8", r"q25_\d", "Support for the EU becoming the main authority in policy areas (education, health, fiscal, military, housing, digital, gender equality, LGBTIQ+ rights)", "eu_integration;gender_policy;lgbtiq", n=0)
battery("q26_1–q26_5", r"q26_\d", "Support for EU democratic reforms (direct election of Commission president, transnational parties, national parliament veto, EU referendums, population-weighted decisions)", "eu_reform;eu_integration")
battery("q27_1–q27_2", r"q27_\d", "Attachment to own country and to the EU", "identity_attachment")
single("q29", "Zero-sum perceptions of EU economic integration (own country poorer vs everyone benefited)", "zero_sum;eu_integration", n=0)
battery("q31_1–q31_5", r"q31_\d", "Regime preferences: strong leader, direct democracy, businesspeople, experts, army rule", "regime_preferences")
single("group", "Random split-ballot assignment (group 1: 0–10 trust scales; group 2: −5/+5 distrust–trust scales)", "vignette")
add("express2", "q32a / q32b", "Generalised social trust — split ballot: 0–10 scale (q32a) vs −5/+5 distrust–trust scale (q32b)", lab["q32a"], "0–10 or −5…+5 depending on split", "social_trust;vignette")
battery("q33a_1–q33a_5, q33b_1–q33b_5", r"q33a_\d", "Trust in the European Parliament, European Commission, national parliament, EU-level politicians and national politicians — split ballot (0–10 trust vs −5/+5 distrust–trust)", "trust_eu;trust_political;vignette", n=0)
battery("q34_1–q34_9", r"q34_\d", "Trust and distrust in the EU battery (competence, integrity, benevolence, scepticism)", "trust_eu")
add("express2", "c1–c5, q35_*", "Conjoint experiment (5 tasks): who designs EU public policy (incl. AI), policy topic, who supports it; choice, trust in the process (0–10) and likelihood of following news",
    "Attributes — Designer: " + vals["c1_a1"].replace("\n", "; ") + ". Topic: " + vals["c1_a2"].replace("\n", "; ") + ". Support: " + vals["c1_a3"].replace("\n", "; ") + ". Outcomes: " + lab["q35_1_2"] + " / " + lab["q35_1_3"],
    "Forced choice A/B; 0–10 ratings", "conjoint;ai_governance;trust_eu")
battery("q36_1–q36_7", r"q36_\d", "Trust in digital technologies (search engines, social media, generative AI, facial recognition, ranking algorithms, public-sector algorithms)", "digital_rights;ai_attitudes")
battery("q37_1–q37_2", r"q37_\d", "Perceived political corruption in own country and in the EU", "corruption;eu_integration")
battery("q39_1–q39_10", r"q39_\d+", "Gender equality attitudes: EU's role, backlash ('gone too far', quotas unfair), misinformation, voting for women candidates, women's capability for office", "gender_attitudes;gender_backlash;gender_policy;women_representation;eu_integration")
single("q40", "Importance of complying with international agreements (Istanbul Convention) on gender-based violence", "gbv", n=0)
battery("q41_1–q41_13", r"q41_\d+", "Perceived root causes of gender-based violence (select 3)", "gbv", n=0)
battery("q42_1–q42_8", r"q42_\d", "Who is responsible for preventing gender-based violence (select up to 2)", "gbv", n=0)
battery("q43_1–q43_2", r"q43_\d", "Traditional family values vs LGBTIQ+ rights and vs women's rights", "traditional_values;lgbtiq;gender_attitudes")
battery("q45_1–q45_8", r"q45_\d", "Perceived impact of digitalisation on rights, public services, inclusion and the economy", "digital_rights")
battery("q46_1–q46_7", r"q46_\d", "Actions taken to protect one's rights online", "digital_rights", n=0)
battery("q47_1–q47_10", r"q47_\d+", "Concerns and beliefs about AI and digital technologies (job losses, discrimination, disinformation, big tech power, EU digital policy, democratic participation)", "ai_attitudes;digital_rights")
battery("q48_1–q48_8", r"q48_\d", "Preferences on AI governance (transparency, explainability, safeguards, environmental impact, EU regulation standards)", "ai_governance")
battery("q49_1–q49_7", r"q49_\d", "Preferred level of human vs AI involvement by domain (courts, legal, public administration, law-making, health, education, transport)", "ai_governance;ai_attitudes", n=0)
add("express2", "q50rank1–q50rank3", "Preferred ways for citizens to contribute to AI regulation (ranking)", lab["q50rank1"], scale("q50rank1", 0), "ai_governance;protest")
battery("q51_1–q51_8", r"q51_\d", "Digital protections citizens want developed at EU level", "digital_rights;eu_integration", n=0)
single("weight", "Post-stratification weight (gender, age, region)", "weights")

# ---- Democracy in the Age of AI (Spanish questionnaire, hand-coded) ----
A = "aidem"
add(A, "P1", "Gender", "¿Eres…?", "Hombre / Mujer / Otro / Prefiero no decirlo", "gender")
add(A, "P2", "Age (years)", "¿Cuántos años tienes?", "Numeric", "age")
add(A, "P3", "Country of birth", "¿En qué país naciste?", "Spain / other (list)", "ethnicity")
add(A, "P4", "Postcode of residence", "¿Podrías indicarme tu actual código postal?", "5-digit postcode", "residence")
add(A, "P6", "Highest level of education", "¿Cuál es el mayor nivel de estudios que has alcanzado?", "8 categories (no studies – PhD)", "education")
add(A, "OCUPACIÓN", "Occupation (ISCO-based), incl. unpaid domestic work, student, retired", "¿Cuál de las siguientes categorías explica mejor tu ocupación?", "14 categories", "occupation;employment")
add(A, "P7, P8", "Number of siblings, and each sibling's sex and relative age", "¿Cuántos hermanos/as tienes? / ¿Podrías indicar la edad y el sexo de tus hermanos/as?", "Count (0–12); sex and older/younger/same age per sibling", "siblings")
add(A, "P9", "Father's and mother's occupation when respondent was ~14", "Pensando en el trabajo principal que tenía su padre/madre cuando tenías aproximadamente 14 años…", "14 categories, separately for father and mother", "social_origin;occupation")
add(A, "HIJOS", "Number of children", "¿Cuántos hijos/as tienes en la actualidad?", "Count (0–10)", "children")
add(A, "P10_1–P10_10", "Personality traits (openness, conscientiousness, extraversion, agreeableness, neuroticism — 10 items)", "¿Hasta qué punto crees que te definen a ti como persona? (interesarme por cosas nuevas, prefiero la rutina, planificar tareas, centro de atención, ponerme en el lugar de los demás, evito conflictos, mantengo la calma, me siento tenso/a…)", "1–5 agreement", "personality")
add(A, "P11_1–P11_4", "Technology-related activities during adolescence (computers/consoles, programming, apps, fixing devices)", "Durante tu adolescencia, ¿con qué frecuencia realizabas las siguientes actividades relacionadas con tecnología?", "1 Never – 5 Very often", "socialisation;social_media")
add(A, "P12", "Type of school attended (public / private / semi-private)", "El centro educativo en el que cursaste la mayor parte de tu educación obligatoria era…", "Público / Privado / Concertado", "socialisation;education")
add(A, "P13", "Ever attended a single-sex school", "¿Alguna vez asististe a un centro de educación diferenciada por sexo?", "Yes / No", "socialisation;gender")
add(A, "P14_1–P14_6", "Gender composition of peers in adolescent activities (games, video games, reading, science, sport, drawing)", "Durante tu adolescencia, ¿con quién solías realizar las siguientes actividades?", "Mainly boys / mainly girls / both / alone / didn't do it", "socialisation;gender_identity")
add(A, "P15, P16", "Life stage when closest female and male friendships began, and where friends were met", "¿En qué etapa de tu vida se originaron tus amistades cercanas? / ¿En qué ámbito conociste a tus amistades?", "Multiple choice, separately for female and male friends", "socialisation")
add(A, "P17, P18", "Self-rated masculinity and femininity (1–10)", "¿Cómo de masculino/a te sientes? / ¿Cómo de femenino/a te sientes?", "1–10", "gender_identity")
add(A, "P19_1–P19_10", "Frequency of leisure and social activities (socialising, going out, sport, culture, reading, associations, volunteering, board games, caring for others)", "¿Prácticas alguna de las siguientes actividades? (… Cuidar y asistir a otras personas de tu ámbito familiar o cercano)", "1 Never – 5 Very often", "leisure;care;protest")
add(A, "P20", "Three cultural products consumed and liked in the last month (open)", "¿Podrías indicar el nombre de tres series, películas, libros, programas, grupos de música, videojuegos…?", "Open text", "leisure")
add(A, "P21_1–P21_12", "Frequency of digital activities (social media, streaming, YouTube/Twitch, podcasts, online news, video games, forums, online shopping, AI tools, pornography)", "¿Con qué frecuencia realizas las siguientes actividades digitales en tu vida cotidiana?", "1 Never – 5 Very often", "social_media;news_media;ai_use")
add(A, "P22", "Daily time spent on social media", "En un día normal, ¿cuánto tiempo dedicas a utilizar redes sociales?", "6 categories (none – 3h+)", "social_media")
add(A, "P23–P25", "Influencers/streamers followed (open), how often, and reasons for following them", "Indica el nombre de hasta tres influencers, streamers o creadores de contenido que sigas… ¿Con qué frecuencia…? ¿Por qué motivos…?", "Open text; frequency; multiple choice (up to 3)", "influencers;social_media")
add(A, "P26", "Parasocial relationship with followed content creators (feel I know them, trust their opinions, identify with them…)", "Pensando en los creadores de contenidos que sigues con más frecuencia, ¿en qué medida estás de acuerdo…?", "1–5 agreement", "influencers")
add(A, "P27", "Relationship with 10 well-known Spanish influencers (Ibai Llanos, Dulceida, María Pombo, El Rubius, El Xokas, Roro…)", "Indica cuál es tu relación con su contenido.", "6 categories (don't know them – watch daily)", "influencers")
add(A, "P28", "Political experiences on social media (unfollowing over politics, rethinking an issue due to a creator, commenting, sharing, fatigue)", "¿Cuántas de las siguientes situaciones te ha ocurrido en los últimos meses en redes sociales?", "Multiple choice", "social_media;influencers;protest")
add(A, "P29", "Interest in politics", "¿En qué medida te interesa la política?", "Nada / Poco / Bastante / Mucho", "political_interest")
add(A, "P29A", "Support for higher taxes on very high incomes", "¿Hasta qué punto estás de acuerdo o en desacuerdo con que las personas con ingresos muy altos deban pagar más impuestos?", "1–5 agreement", "taxation")
add(A, "P30", "Left–right self-placement (0–10)", "En política se suele hablar de izquierda y derecha…", "0 Left – 10 Right", "left_right")
add(A, "P31, P31A–C", "Conjoint experiment: who drafts a policy (politicians, experts, experts + AI, AI guided by politicians, AI alone), who approves it (Government, Congress, EU, AI), problem addressed and policy area; choice, rating and perceived legitimacy",
    "A continuación, te mostraremos dos posibles propuestas de política pública en España… ¿Cuál prefieres? / ¿Cómo valorarías cada uno? / ¿Hasta qué punto consideras válido cada uno de estos procedimientos?", "Forced choice; 1–10 rating; 4-point validity", "conjoint;ai_governance;regime_preferences")
add(A, "P32_1–P32_7", "Frequency of use of generative AI tools (ChatGPT, Copilot, Gemini, Claude, Perplexity, Consensus, Grok)", "¿Con qué frecuencia utilizas las siguientes herramientas de inteligencia artificial generativa?", "6 categories (don't know it – daily)", "ai_use")
add(A, "P33, P34", "Expected impact of generative AI on one's work/studies and on personal life", "¿Cómo crees que la inteligencia artificial generativa cambiará tu trabajo o estudios / tu vida personal?", "1 Very negatively – 5 Very positively", "ai_attitudes;employment")
add(A, "P35_1–P35_8", "Emotions towards generative AI (enthusiasm, interest, trust, worry, rejection, scepticism, uncertainty, loss of control)", "¿Hasta qué punto la inteligencia artificial generativa te despierta los siguientes sentimientos?", "1 Not at all – 4 A lot", "ai_attitudes")
add(A, "P36_1–P36_7", "Support for using AI in public administration (answering questions, simplifying language, automating paperwork, fraud detection, proposing resolutions, recidivism risk, automatic sanctions)", "¿Estarías de acuerdo con la inclusión y uso de inteligencia artificial en la Administración Pública para las siguientes tareas?", "1–5 agreement", "ai_governance")
add(A, "P37, P38", "Main advantages and risks of generative AI for Spanish politics and democracy", "¿Cuáles crees que podrían ser las principales ventajas / desventajas o riesgos del uso de la IA generativa en la política y democracia?", "Multiple choice", "ai_attitudes;ai_governance")
add(A, "P39_1–P39_6", "Regime preferences: strong leader, direct democracy, businesspeople, independent experts, armed forces, representative democracy", "A continuación se presentan distintas formas de organizar el gobierno… ¿Estás a favor o en contra?", "1 Totally against – 5 Totally in favour", "regime_preferences")
add(A, "P40_1–P40_5", "Trust in the Prime Minister, Congress, regional president, mayor and the EU (0–10)", "En general, ¿cuánta confianza tienes en los siguientes actores políticos?", "0 No trust – 10 Total trust", "trust_political;trust_eu")
add(A, "P41, P41A–D", "Vignette experiment: a person who won €1m/€5m/€50m as a TV quiz winner, lottery winner, footballer, entrepreneur, crypto investor or influencer — fair tax rate, deservingness, paying half in tax, moving abroad to avoid tax",
    "Imagina a una persona que vive en España y que recientemente ha ganado [Cantidad]. Esta persona es [Perfil]…", "Slider 0–100%; 4- and 5-point scales", "vignette;taxation;deservingness")
add(A, "P42_1–P42_8", "Sexism and perceived gender equality (men priority for jobs, women exaggerate, men should protect women, sexist media, equal opportunities…)", "Pensando ahora en la situación de hombres y mujeres en nuestra sociedad hoy en día…", "1–5 agreement", "gender_attitudes;gender_backlash")
add(A, "P43_1–P43_9", "Attitudes towards feminism (necessary, over-politicised, divides society, as dangerous as machismo, promotes hatred of men, gone too far…)", "¿Hasta qué punto estás de acuerdo con las siguientes afirmaciones? (El feminismo es necesario… ha llegado demasiado lejos)", "1–5 agreement", "feminism;gender_backlash")
add(A, "P44A", "Preferred regulation of sex work (abolition to decriminalisation)", "Sobre el trabajo sexual…", "5 policy options", "morality_policy")
add(A, "P44B", "Preferred regulation of surrogacy", "Sobre la gestación subrogada…", "6 policy options", "morality_policy")
add(A, "P44C", "Preferred abortion regulation", "Sobre el aborto…", "6 policy options", "abortion")
add(A, "P44D", "Preferred regulation of trans rights / gender self-determination", "Sobre las personas trans…", "6 policy options", "lgbtiq")
add(A, "P45", "Recalled vote in the 2023 Spanish general election (incl. abstention)", "¿A qué partido votaste en las anteriores elecciones generales de 2023?", "Party list + blank/null/abstained", "vote_choice;turnout")
add(A, "P46", "Propensity to vote for each party (PP, PSOE, Vox, Sumar, SALF, Podemos, regional parties)", "Si se celebraran elecciones generales mañana, ¿qué probabilidad tendrías de votar por los siguientes partidos?", "1–10", "ptv;vote_choice")
add(A, "P47", "Marital status", "¿Cuál es tu estado civil?", "7 categories", "marital")
add(A, "P48", "Sexual orientation", "Dirías que tu orientación sexual es…", "7 categories", "sexual_orientation")
add(A, "P49–P51", "Current relationship situation, desire for a stable partner, and main reason for not having one (incl. ideological differences, 'feminist radicalisation')", "¿Cuál de las siguientes describe mejor tu situación sentimental actual? / ¿En qué medida te gustaría tener una pareja estable? / ¿Cuál dirías que es el principal motivo…?", "Categorical; 1–5", "marital;gender_backlash")

# ---- Twice as Hard, Half as Good? (Spanish questionnaire, hand-coded) ----
T = "twiceasgood"
add(T, "P1", "Gender", "Género: ¿Cuál de las siguientes opciones te describe mejor?", "Hombre / Mujer / Otro / Prefiero no contestar", "gender")
add(T, "P2", "Age (years)", "¿Cuántos años tienes?", "Numeric", "age")
add(T, "P3", "Autonomous community of residence", "¿En qué Comunidad Autónoma resides actualmente?", "19 regions", "residence")
add(T, "P4", "Highest level of education", "¿Cuál es el mayor nivel de educación que ha completado?", "8 categories", "education")
add(T, "P5", "Interest in politics", "¿Hasta qué punto dirías que estás interesado en política?", "4-point scale", "political_interest")
add(T, "P6_1–P6_6", "How often the 2023 campaign was followed through TV, radio, print press, online news, social media, and sharing on social media", "Desde que se anunciaron [las elecciones], ¿con qué frecuencia has seguido la campaña a través de las siguientes actividades?", "1 Not at all – 5 Every day", "campaign;news_media;social_media")
add(T, "P7a", "Watched the 2023 leaders' debates", "¿Has visto alguno de estos debates?", "All / some / only Sánchez–Feijóo / none", "campaign")
add(T, "P7b", "Perceived tone of each leader in the debates (Sánchez, Feijóo, Abascal, Díaz) on 5 dimensions: rude–polite, heated–calm, emotional–cold, hostile–friendly, confrontational–diplomatic",
    "¿Qué opinas del tono de los debates entre los líderes…? (Descortés–Educado, Acalorado–Calmado, Emotivo–Frío, Hostil–Amistoso, Conflictivo–Diplomático)", "1–5 semantic differentials per leader", "leader_eval;campaign;incivility;women_representation")
add(T, "P8", "Contacted by parties (PSOE, PP, Vox, Sumar) during the campaign, by channel (email, social media, phone, mail, face to face)", "¿Se han puesto en contacto contigo alguno de los siguientes partidos…?", "Multiple choice grid", "campaign")
add(T, "P9", "Turnout in the July 2023 general election (in person / by post / no)", "¿Has votado en estas elecciones de Julio de 2023?", "Categorical", "turnout")
add(T, "P10a, P10b", "Vote choice in the 2023 general election — Congress and Senate", "¿A qué partido votaste… al Congreso de los Diputados / en el Senado?", "Party list + blank/null", "vote_choice")
add(T, "P11a", "Like–dislike of 8 parties (0–10)", "¿Cuál es tu opinión sobre los siguientes partidos políticos? (0–10)", "0 Dislike a lot – 10 Like a lot", "ptv")
add(T, "P11b", "Like–dislike of 8 party leaders (0–10), incl. Yolanda Díaz and Míriam Nogueras", "¿Cuánto te gusta [Nombre del líder político]?", "0 Dislike a lot – 10 Like a lot", "leader_eval;women_representation")
add(T, "P12", "Recalled vote in the May 2023 local and regional elections", "¿Recuerdas a qué partidos votaste en estas elecciones [locales y regionales]?", "Party list", "vote_choice;turnout")
add(T, "P13", "Recalled vote in the 2019 general election", "¿Y recuerdas a qué partido votaste en las anteriores elecciones generales celebradas en 2019?", "Party list", "vote_choice;turnout")
add(T, "P14", "Evaluation of the Sánchez government's performance", "¿En qué medida crees que el presidente Pedro Sánchez y su gobierno han hecho un buen o mal trabajo…?", "4-point scale", "gov_eval")
add(T, "P15", "Retrospective evaluation of the national economy (last 12 months)", "¿Dirías que, en los últimos doce meses, la situación de la economía en España ha mejorado…?", "5-point scale", "economic_perceptions")
add(T, "P16, P16a", "Personal/household economic situation and change over the last year", "¿Cómo calificarías tu situación económica personal (en tu hogar)? / ¿…es mejor, igual o peor que hace un año?", "5-point; 3-point", "subjective_economy")
add(T, "P17a, P17b", "Perceived change in unemployment and prices over the last year", "¿Cuál de las siguientes frases sobre el paro y los precios es la más acertada?", "Increased / same / decreased", "economic_perceptions")
add(T, "P18a, P18b", "Self-rated masculinity and femininity (0–6)", "¿Cómo de masculino dirías que te sientes? / ¿Cómo de femenina dirías que te sientes?", "0–6", "gender_identity")
add(T, "P19a–c", "Importance of being a man/woman (or one's gender) for self-definition", "¿Qué importancia tiene ser hombre/mujer para definirte a ti mismo/a?", "0–6", "gender_identity")
add(T, "P20_1–P20_4", "Gendered personality traits: risk-taking, leadership, sensitivity to others' needs, empathy", "¿Hasta qué punto crees que te definen a ti como persona? (asumir riesgos, ejercer de líder, sensible a las necesidades de los demás, ponerme en el lugar de los demás)", "0–6", "personality;gender_identity")
add(T, "P21", "Personal experience of unfair treatment because of one's gender", "¿Te has sentido alguna vez que se te ha tratado injustamente por tu género?", "Once / several times / no", "discrimination;gender_attitudes")
add(T, "P22–P25", "Perceived gender (in)equality in treatment: health care, employment, standing as a political candidate, policing", "¿Dirías que las mujeres y los hombres reciben el mismo trato justo a la hora de…?", "Women treated worse / men treated worse / same", "discrimination;gender_attitudes;women_representation")
add(T, "P26_1–P26_3", "Support for gender equality measures: parity in parliament, sanctions for sexist comments at work, fines for unequal pay", "¿Hasta qué punto estás a favor o en contra de las siguientes medidas?", "1 Totally in favour – 5 Totally against", "gender_policy;women_representation")
add(T, "P27", "Support for mandatory equal paid parental leave for both parents", "¿Hasta qué punto estarías a favor o en contra de una medida legal que obligue a ambos progenitores a disfrutar de periodos iguales de permiso retribuido…?", "1 Totally in favour – 5 Totally against", "gender_policy;care")
add(T, "P28", "Left–right self-placement (0–10)", "Usando una escala que vaya del 0 al 10, siendo 0 extrema izquierda y 10 extrema derecha…", "0–10", "left_right")
add(T, "P29_1–P29_3", "Support for trans anti-discrimination protections, gender self-ID from 16, and criminalising sex without verbal consent", "¿Estás a favor o en contra de las siguientes medidas?", "1 Totally in favour – 5 Totally against", "lgbtiq;gbv")
add(T, "P30, P31", "Climate change beliefs (has it changed, causes) and expected effects for Spain", "¿Cuál de las siguientes afirmaciones se acerca más a tu opinión [sobre el clima]? / ¿Cómo de malos o buenos cree que serán los efectos…?", "Categorical; 0–10", "climate")
add(T, "P32", "Energy priority: renewables vs fossil fuels", "¿Cuál de las siguientes opciones cree que debería ser la prioridad más importante para asegurar el abastecimiento energético?", "Categorical", "climate")
add(T, "P33_1–P33_4", "Hostile/modern sexism: women seek power by controlling men, are easily offended, are paid less, exaggerate harassment claims", "¿Con qué frecuencia crees que…?", "1 Never – 5 Always", "gender_attitudes;gender_backlash")
add(T, "P34_1–P34_2", "Benevolent sexism: women's superior moral sensitivity; men should protect women", "¿Hasta qué punto estás de acuerdo o en desacuerdo con las siguientes dos afirmaciones?", "1 Totally agree – 5 Totally disagree", "gender_attitudes")
add(T, "P35_1–P35_4", "Masculinity norms (aggression, multiple sexual partners), gender self-ID as a threat to family values, same-sex adoption rights", "Hasta qué punto estás de acuerdo o en desacuerdo con cada una de las siguientes afirmaciones…", "1 Totally agree – 5 Totally disagree", "gender_attitudes;lgbtiq;traditional_values")
add(T, "P36_1–P36_4", "Child-rearing values (authoritarianism scale: independence vs respect, obedience vs self-reliance…)", "A continuación se muestran parejas de cualidades deseables en los niños…", "Binary choice per pair", "traditional_values")
add(T, "P37", "Satisfaction with how democracy works in Spain", "¿Cuánto de satisfecho o insatisfecho dirías que estás con el funcionamiento de la democracia en España?", "4-point scale", "democracy_satisfaction")
add(T, "P38", "Sympathy with the feminist movement", "¿Hasta qué punto simpatizas con el movimiento feminista?", "4-point scale", "feminism")
add(T, "P39", "Participation in 8 March (International Women's Day) mobilisations", "¿Has participado alguna vez en la movilización del 8 de Marzo?", "2023 / earlier / never", "feminism;protest")
add(T, "P40, P41", "Perceived conflictiveness of Spanish politics; need to compromise vs stand firm", "¿Cómo de conflictivo calificarías que es el ambiente político actual…? / en política es necesario ceder… mantenerse firme", "1–5", "incivility")
add(T, "P42_1–P42_6", "Acceptability of uncivil behaviour by elected officials (insults, attacks on appearance/family/identity, interrupting)", "¿Con qué frecuencia dirías que es aceptable que los cargos electos…?", "Always – never acceptable", "incivility")
add(T, "P43, P44", "Justification of political violence/intimidation, and perceived change over four years", "¿Hasta qué punto crees que puede estar justificado que la gente utilice cualquier tipo de violencia…? / …ha incrementado…?", "5-point scales", "political_violence")
add(T, "P45, P46", "Religious identification and attendance at religious services", "¿Con qué religión te identificas? / ¿Con qué frecuencia asiste a servicios religiosos?", "Categorical; 6-point", "religion")
add(T, "Q47", "Employment status", "¿En cuál de las siguientes situaciones te encuentras actualmente?", "9 categories", "employment")
add(T, "P48", "Province where respondent voted", "¿En qué provincia has votado en estas elecciones del 23 de julio?", "52 provinces", "residence")
add(T, "P49", "Type of place of residence (farm to big city)", "¿Cómo describirías el lugar (ciudad o pueblo) donde actualmente resides?", "6 categories", "residence")
add(T, "P50", "Marital status", "¿Cuál es tu estado civil?", "4 categories", "marital")
add(T, "P51–P53", "Has children; age of youngest child; has a daughter", "¿Tienes hijos/as? / ¿Cuál es la edad de tu hijo/a menor? / ¿Alguno/a de tus hijos/as es mujer?", "Categorical", "children")
add(T, "P54", "Sexual orientation", "Dirías que tu orientación sexual es…", "Heterosexual / homosexual / bisexual / other", "sexual_orientation")

# ---------------------------------------------------------------- checks + write
cids = {c[0] for c in CONCEPTS}
tids = {t[0] for t in THEMES}
for c in CONCEPTS:
    for t in c[2].split(";"):
        assert t in tids, (c, t)
for v in V:
    for c in v["concepts"].split(";"):
        assert c in cids, (v["variable"], c)

def write(name, fields, recs):
    with open(os.path.join(OUT, name), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in recs: w.writerow(r)

write("themes.csv", ["theme_id", "theme_label"], [dict(theme_id=a, theme_label=b) for a, b in THEMES])
write("concepts.csv", ["concept_id", "concept_label", "themes", "synonyms"],
      [dict(concept_id=a, concept_label=b, themes=c, synonyms=d) for a, b, c, d in CONCEPTS])
write("surveys.csv", list(SURVEYS[0].keys()), SURVEYS)
write("variables.csv", ["survey_id", "variable", "question_en", "question_original", "response_scale", "concepts"], V)
from collections import Counter
print(len(V), "variables", Counter(v["survey_id"] for v in V))
