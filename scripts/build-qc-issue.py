#!/usr/bin/env python3
"""Quality-control issue of the Omni Law Gazette — 15–22 May 2026.

Compiled from the official Azerbaijani-language portal material (president.az,
nk.gov.az; Constitutional Court and e-qanun unavailable for the window). Per
AGENTS.md, the Azerbaijani text is authored first and faithful to the source
legal wording; the English text is its translation.
"""
import os
from gazette_lib import build_issue

OUTDIR = os.path.join(os.path.dirname(__file__), 'gazettes')
BASENAME = 'Omni-Law-Gazette-QC-2026-05-15-to-22'

CHROME = {
    'AZ': {
        'banner': 'HƏFTƏLİK QANUNVERİCİLİK İCMALI · AZƏRBAYCAN RESPUBLİKASI',
        'issue':  'KEYFİYYƏT YOXLAMASI · 15–22 MAY 2026',
        'footer_lead': 'Omni Law Gazette',
        'footer': '  ·  Azərbaycan Respublikasının qanunvericilik fəaliyyətinin həftəlik daxili icmalı; '
                  'Omni Law Firm tərəfindən yalnız məlumat məqsədilə hazırlanıb. Hüquqi məsləhət təşkil etmir; '
                  'istənilən maddəyə istinad etməzdən əvvəl e-qanun.az saytındakı konsolidə edilmiş mətnlərə baxın. '
                  'Mənbələr: president.az · nk.gov.az · meclis.gov.az · e-qanun.az · constcourt.gov.az',
        'footer_c': 'Keyfiyyət yoxlaması buraxılışı  ·  Baxış dövrü 15–22 may 2026  ·  © 2026 Omni Law Firm',
    },
    'EN': {
        'banner': 'WEEKLY LEGISLATIVE DIGEST · REPUBLIC OF AZERBAIJAN',
        'issue':  'QUALITY-CONTROL EDITION · 15–22 MAY 2026',
        'footer_lead': 'Omni Law Gazette',
        'footer': '  ·  Internal weekly digest of the legislative activity of the Republic of Azerbaijan, '
                  'compiled by Omni Law Firm for informational purposes only. It does not constitute legal advice; '
                  'consult the consolidated texts on e-qanun.az before relying on any item. '
                  'Sources: president.az · nk.gov.az · meclis.gov.az · e-qanun.az · constcourt.gov.az',
        'footer_c': 'Quality-control edition  ·  Reporting period 15–22 May 2026  ·  © 2026 Omni Law Firm',
    },
}

LEAD = {
    'n': '01',
    'kicker': {'AZ': 'PREZİDENT · QANUN VƏ FƏRMAN · 19 MAY 2026',
               'EN': 'PRESIDENT · LAW & DECREE · 19 MAY 2026'},
    'headline': {'AZ': 'Əmək Məcəlləsinə və narkoloji xidmət qanununa dəyişikliklər qüvvəyə minir.',
                 'EN': 'Amendments to the Labour Code and the narcological-service law take effect.'},
    'body': {
        'AZ': [
            "Baxış dövrünün mərkəzi hadisəsi Azərbaycan Respublikasının Əmək Məcəlləsində və “Narkoloji xidmət "
            "və nəzarət haqqında” Qanunda edilmiş dəyişikliklərdir. Milli Məclisin 27 aprel 2026-cı il tarixli "
            "394-VIIQD nömrəli Qanunu ilə qəbul edilmiş düzəlişlər 19 mayda Prezidentin fərmanı ilə tətbiq edilmiş "
            "və müvafiq icra mexanizmləri müəyyən olunmuşdur.",
            "Dəyişikliklər əmək münasibətlərini narkoloji nəzarət tələbləri ilə əlaqələndirir və işəgötürənlər üçün "
            "yeni öhdəliklər yarada bilər. Əmək hüququ üzrə praktiklər düzəlişlərin tam mətnini və qüvvəyəminmə "
            "şərtlərini diqqətlə nəzərdən keçirməlidirlər.",
        ],
        'EN': [
            "The central development of the review period is a set of amendments to the Labour Code of the Republic "
            "of Azerbaijan and to the Law on Narcological Service and Control. Adopted by the Milli Majlis as Law "
            "No. 394-VIIQD of 27 April 2026, the changes were brought into application by presidential decree on "
            "19 May, which set out the corresponding implementing mechanisms.",
            "The amendments link employment relations to narcological-control requirements and may create new "
            "obligations for employers. Employment-law practitioners should review the full text of the amendments "
            "and their entry-into-force conditions carefully.",
        ],
    },
            'eqanun': [61868],
}

ARTICLES = [
    {
        'n': '02',
        'kicker': {'AZ': 'PREZİDENT · QANUNLAR · 19 MAY 2026', 'EN': 'PRESIDENT · LAWS · 19 MAY 2026'},
        'headline': {'AZ': 'Hüquqi yardım sahəsində üç beynəlxalq sənəd təsdiq edilir.',
                     'EN': 'Three international legal-assistance instruments are ratified.'},
        'body': {
            'AZ': [
                "Prezident məhkəmə və digər səlahiyyətli orqanların aktlarının məcburi icrası zamanı məlumat "
                "mübadiləsi üzrə əməkdaşlıq haqqında Sazişi təsdiq edən Qanunu, habelə mülki, ailə və cinayət işləri "
                "üzrə hüquqi yardım və hüquqi münasibətlər haqqında 1993-cü il və 2002-ci il Konvensiyalarına "
                "dəyişiklik edən iki Protokolu təsdiq edən qanunları imzalamışdır.",
                "Bu sənədlər sərhəddənkənar icra və hüquqi yardım mexanizmlərini gücləndirir. Beynəlxalq mübahisələr, "
                "habelə qərarların tanınması və icrası ilə məşğul olan praktiklər üçün əhəmiyyət daşıyır.",
            ],
            'EN': [
                "The President signed a law approving the Agreement on cooperation in the exchange of information "
                "during the compulsory enforcement of acts of courts and other competent bodies and officials, "
                "together with laws approving two Protocols amending the 1993 and 2002 Conventions on legal "
                "assistance and legal relations in civil, family and criminal matters.",
                "These instruments strengthen the mechanisms for cross-border enforcement and legal assistance. They "
                "are of relevance to practitioners engaged in international disputes and in the recognition and "
                "enforcement of judgments.",
            ],
        },
                'eqanun': [61863, 61860, 61864],
    },
    {
        'n': '03',
        'kicker': {'AZ': 'PREZİDENT · FƏRMAN · 20 MAY 2026', 'EN': 'PRESIDENT · DECREE · 20 MAY 2026'},
        'headline': {'AZ': 'Xüsusi rabitə xidmətlərinin tarif tənzimlənməsi təkmilləşdirilir.',
                     'EN': 'Tariff regulation for special-communications services is refined.'},
        'body': {
            'AZ': [
                "20 may tarixli fərman Xüsusi Rabitə və İnformasiya Təhlükəsizliyi Dövlət Xidməti tərəfindən "
                "göstərilən xidmətlərin tariflərinin tənzimlənməsini təkmilləşdirir və “Feldyeger rabitəsi haqqında” "
                "Qanunun tətbiqini təmin edir.",
                "Sənəd dövlət rabitə və informasiya təhlükəsizliyi xidmətlərinin tarif rejimini yeniləyir. Bu sahədə "
                "fəaliyyət göstərən və ya bu xidmətlərdən istifadə edən qurumlar yeni tarif qaydalarını nəzərə "
                "almalıdırlar.",
            ],
            'EN': [
                "A decree of 20 May refines the regulation of the tariffs for services provided by the Special "
                "Communications and Information Security State Service and gives effect to the Law on Courier "
                "(Feldjäger) Communications.",
                "The instrument updates the tariff regime for state communications and information-security services. "
                "Bodies operating in, or relying on, this sector should take account of the new tariff rules.",
            ],
        },
                'eqanun': [61871],
    },
    {
        'n': '04',
        'kicker': {'AZ': 'PREZİDENT · FƏRMAN · 19 MAY 2026', 'EN': 'PRESIDENT · DECREE · 19 MAY 2026'},
        'headline': {'AZ': 'Bolqarıstanda “Şuşa Parkı”nın salınmasına dair Saziş təsdiqlənir.',
                     'EN': 'An agreement on a “Shusha Park” in Bulgaria is approved.'},
        'body': {
            'AZ': [
                "Prezident Azərbaycan Respublikası Hökuməti ilə Bolqarıstan Respublikası Hökuməti və Veliko Tırnovo "
                "Bələdiyyəsi arasında Veliko Tırnovo şəhərində “Şuşa Parkı”nın salınması və yenidən qurulması üzrə "
                "əməkdaşlıq haqqında Sazişi təsdiq edən fərman imzalamışdır.",
                "Saziş mədəni diplomatiya çərçivəsində ikitərəfli əməkdaşlığı möhkəmləndirir və Şuşanın mədəni "
                "irsinin xaricdə təmsil olunmasını nəzərdə tutur.",
            ],
            'EN': [
                "The President signed a decree approving the Agreement — between the Government of Azerbaijan on the "
                "one hand and the Government of Bulgaria and the Municipality of Veliko Tarnovo on the other — on "
                "cooperation in laying out and reconstructing a “Shusha Park” in the city of Veliko Tarnovo.",
                "The agreement reinforces bilateral cooperation as a matter of cultural diplomacy and provides for "
                "the representation of Shusha’s cultural heritage abroad.",
            ],
        },
                'eqanun': [61862],
    },
    {
        'n': '05',
        'kicker': {'AZ': 'PREZİDENT · SƏRƏNCAM · 20 MAY 2026', 'EN': 'PRESIDENT · ORDER · 20 MAY 2026'},
        'headline': {'AZ': '2026-cı ilin iyul çağırışı elan olunur.',
                     'EN': 'The July 2026 conscription call-up is announced.'},
        'body': {
            'AZ': [
                "Prezidentin 20 may tarixli sərəncamı ilə Azərbaycan Respublikası vətəndaşlarının 2026-cı il iyulun "
                "1-dən 30-dək müddətli həqiqi hərbi xidmətə çağırılması, eyni zamanda müddətli həqiqi hərbi xidmət "
                "hərbi qulluqçularının ehtiyata buraxılması müəyyən edilmişdir.",
                "Sərəncam növbəti çağırış kampaniyasının vaxtını və əhatəsini təsbit edir. Çağırış yaşında olan "
                "vətəndaşlar və işəgötürənlər müvafiq müddətləri nəzərə almalıdırlar.",
            ],
            'EN': [
                "By a presidential order of 20 May, citizens of the Republic of Azerbaijan are called up for "
                "fixed-term active military service from 1 to 30 July 2026, and serving fixed-term conscripts are to "
                "be discharged to the reserve.",
                "The order fixes the timing and scope of the next call-up campaign. Citizens of conscription age and "
                "their employers should take note of the relevant dates.",
            ],
        },
                'eqanun': [61870],
    },
    {
        'n': '06',
        'kicker': {'AZ': 'NAZİRLƏR KABİNETİ · QƏRAR № 159 · 21 MAY 2026',
                   'EN': 'CABINET · DECISION No. 159 · 21 MAY 2026'},
        'headline': {'AZ': 'Vahid xəzinə hesabının qalığının idarə edilməsi qaydaları təsdiqlənir.',
                     'EN': 'Rules adopted for managing the single treasury account balance.'},
        'body': {
            'AZ': [
                "Nazirlər Kabineti “Vahid xəzinə hesabının qalığının idarəetməyə verilməsi Qaydaları”nı təsdiq "
                "etmişdir. Sənəd dövlət vəsaitlərinin vahid xəzinə hesabındakı qalığının idarə edilməsi mexanizmini "
                "müəyyən edir.",
                "Qaydalar dövlət maliyyəsinin idarə olunmasında mühüm addımdır və xəzinə əməliyyatlarının "
                "şəffaflığına təsir göstərir.",
            ],
            'EN': [
                "The Cabinet of Ministers approved the Rules on transferring the balance of the single treasury "
                "account to management. The instrument establishes the mechanism for managing the balance of state "
                "funds held in the single treasury account.",
                "The Rules are a significant step in public-finance management and bear on the transparency of "
                "treasury operations.",
            ],
        },
                'eqanun': [61887],
    },
    {
        'n': '07',
        'kicker': {'AZ': 'NAZİRLƏR KABİNETİ · SƏRƏNCAM № 326s · 15 MAY 2026',
                   'EN': 'CABINET · ORDER No. 326s · 15 MAY 2026'},
        'headline': {'AZ': '“Bakı Metropoliteni”nin yenidən təşkili icra olunur.',
                     'EN': 'The reorganisation of the Baku Metro is implemented.'},
        'body': {
            'AZ': [
                "Nazirlər Kabineti “Bakı Metropoliteni” Qapalı Səhmdar Cəmiyyətinin yenidən təşkilinə dair "
                "Prezidentin 7 may 2026-cı il tarixli 662 nömrəli Fərmanının icrasını təmin edən sərəncam vermişdir; "
                "bununla bağlı bəzi prezident aktlarında dəyişikliklər edilmiş və bəzi sərəncamlar ləğv edilmişdir.",
                "Yenidən təşkil paytaxt metro şəbəkəsinin idarəetmə strukturuna təsir göstərir. Şirkətlə müqavilə "
                "münasibətlərində olan tərəflər dəyişiklikləri izləməlidirlər.",
            ],
            'EN': [
                "The Cabinet of Ministers issued an order giving effect to Presidential Decree No. 662 of 7 May 2026 "
                "on the reorganisation of “Baku Metro” Closed Joint-Stock Company; related amendments were made to "
                "certain presidential acts and certain orders were repealed.",
                "The reorganisation affects the management structure of the capital’s metro network. Parties in "
                "contractual relations with the company should monitor the changes.",
            ],
        },
                'eqanun': [61942],
    },
    {
        'n': '08',
        'kicker': {'AZ': 'NAZİRLƏR KABİNETİ · QƏRAR № 158 · 21 MAY 2026',
                   'EN': 'CABINET · DECISION No. 158 · 21 MAY 2026'},
        'headline': {'AZ': 'Peşə əmək qabiliyyətinin itirilməsi dərəcəsinin müəyyən edilməsi qaydaları təsdiqlənir.',
                     'EN': 'Rules set for assessing loss of occupational working capacity.'},
        'body': {
            'AZ': [
                "Nazirlər Kabineti istehsalatda bədbəxt hadisələr və peşə xəstəlikləri nəticəsində zərər çəkmiş "
                "sığortalının peşə əmək qabiliyyətinin itirilməsi dərəcəsinin müəyyən edilməsi Qaydasını təsdiq "
                "etmişdir.",
                "Qaydalar icbari sığorta üzrə ödənişlərin müəyyənləşdirilməsində istinad çərçivəsi yaradır və "
                "işçilərin sosial müdafiəsinə təsir göstərir.",
            ],
            'EN': [
                "The Cabinet of Ministers approved the Rules for determining the degree of loss of occupational "
                "working capacity of an insured person harmed by industrial accidents and occupational diseases.",
                "The Rules provide the reference framework for determining compulsory-insurance payments and bear on "
                "the social protection of employees.",
            ],
        },
                'eqanun': [61888],
    },
    {
        'n': '09',
        'kicker': {'AZ': 'NAZİRLƏR KABİNETİ · QƏRAR № 160 · 22 MAY 2026',
                   'EN': 'CABINET · DECISION No. 160 · 22 MAY 2026'},
        'headline': {'AZ': '“Əlilliyin qiymətləndirilməsi Qaydası”na dəyişiklik edilir.',
                     'EN': 'The Rules on disability assessment are amended.'},
        'body': {
            'AZ': [
                "Nazirlər Kabineti 2020-ci il 16 yanvar tarixli 11 nömrəli Qərarı ilə təsdiq edilmiş “Əlilliyin "
                "qiymətləndirilməsi Qaydası”na dəyişiklik etmişdir.",
                "Dəyişiklik əlilliyin qiymətləndirilməsi prosedurlarına təsir göstərir və sosial təminat hüquqlarının "
                "müəyyənləşdirilməsində əhəmiyyət daşıyır.",
            ],
            'EN': [
                "The Cabinet of Ministers amended the Rules on the Assessment of Disability approved by its Decision "
                "No. 11 of 16 January 2020.",
                "The amendment affects the procedures for assessing disability and is relevant to the determination "
                "of social-security entitlements.",
            ],
        },
                'eqanun': [61889],
    },
    {
        'n': '10',
        'kicker': {'AZ': 'NAZİRLƏR KABİNETİ · QƏRAR № 157 · 21 MAY 2026',
                   'EN': 'CABINET · DECISION No. 157 · 21 MAY 2026'},
        'headline': {'AZ': 'İşğaldan azad olunmuş ərazilərdə çalışan əməkdaşlara mənzil kirayəsi ödənişi.',
                     'EN': 'Housing-rent reimbursement for staff in the liberated territories.'},
        'body': {
            'AZ': [
                "Nazirlər Kabineti dövlət orqanlarının işğaldan azad edilmiş ərazilərdə işləyən və ya xidmətdə olan "
                "əməkdaşları üçün mənzil kirayəsi xərclərinin ödənilməsini və bununla bağlı bəzi məsələləri "
                "tənzimləyən qərar qəbul etmişdir.",
                "Qərar bərpa və yenidənqurma prosesində kadrların təminatına yönəlmişdir.",
            ],
            'EN': [
                "The Cabinet of Ministers adopted a decision regulating the reimbursement of housing-rent expenses "
                "for employees of state bodies who work or serve in the liberated territories, together with certain "
                "related matters.",
                "The decision is directed at staffing the reconstruction and rehabilitation effort.",
            ],
        },
                'eqanun': [61898],
    },
    {
        'n': '11',
        'kicker': {'AZ': 'NAZİRLƏR KABİNETİ · QƏRAR № 154 · 20 MAY 2026',
                   'EN': 'CABINET · DECISION No. 154 · 20 MAY 2026'},
        'headline': {'AZ': 'Dəmir yolu ilə sərnişin giriş-çıxışına dair qaydalar müəyyən edilir.',
                     'EN': 'Rules set for rail passenger entry and exit.'},
        'body': {
            'AZ': [
                "Nazirlər Kabineti dəmiryol nəqliyyatı vasitəsilə sərnişinlərin Azərbaycan Respublikasına giriş-çıxışı "
                "ilə bağlı bəzi məsələləri tənzimləyən qərar qəbul etmişdir.",
                "Qərar sərhəd keçidi və beynəlxalq dəmir yolu sərnişin daşımaları üçün əhəmiyyət daşıyır.",
            ],
            'EN': [
                "The Cabinet of Ministers adopted a decision regulating certain matters concerning the entry to and "
                "exit from the Republic of Azerbaijan of passengers by rail transport.",
                "The decision is of relevance to border crossing and to international rail passenger transport.",
            ],
        },
                'eqanun': [61878],
    },
    {
        'n': '12',
        'kicker': {'AZ': 'NAZİRLƏR KABİNETİ · SƏRƏNCAM № 349s · 21 MAY 2026',
                   'EN': 'CABINET · ORDER No. 349s · 21 MAY 2026'},
        'headline': {'AZ': 'Tullantıların idarə edilməsi üzrə əlavə tədbirlər icra olunur.',
                     'EN': 'Additional waste-management measures are implemented.'},
        'body': {
            'AZ': [
                "Nazirlər Kabineti tullantıların idarə edilməsinin təkmilləşdirilməsi sahəsində əlavə tədbirlərə dair "
                "Prezidentin 14 may 2026-cı il tarixli 665 nömrəli Fərmanının icrasını təmin edən sərəncam vermişdir.",
                "Tədbirlər ekoloji tənzimləmə çərçivəsini gücləndirir və tullantı sahəsində fəaliyyət göstərən "
                "subyektlərə təsir göstərə bilər.",
            ],
            'EN': [
                "The Cabinet of Ministers issued an order giving effect to Presidential Decree No. 665 of 14 May 2026 "
                "on additional measures to improve waste management.",
                "The measures strengthen the environmental-regulation framework and may affect entities operating in "
                "the waste sector.",
            ],
        },
                'eqanun': [61976],
    },
    {
        'n': '13',
        'kicker': {'AZ': 'NAZİRLƏR KABİNETİ · QƏRAR № 149 · 15 MAY 2026',
                   'EN': 'CABINET · DECISION No. 149 · 15 MAY 2026'},
        'headline': {'AZ': 'Medianın İnkişafı Agentliyinin yenidən təşkili tənzimlənir.',
                     'EN': 'The reorganisation of the Media Development Agency is regulated.'},
        'body': {
            'AZ': [
                "Nazirlər Kabineti “Azərbaycan Respublikasının Medianın İnkişafı Agentliyi” və “Beynəlxalq "
                "Münasibətlərin Təhlili Mərkəzi” publik hüquqi şəxslərinin yenidən təşkilinə dair Prezidentin "
                "20 fevral 2026-cı il tarixli 601 nömrəli Fərmanının icrası ilə əlaqədar öz qərarlarına dəyişikliklər "
                "etmişdir.",
                "Dəyişikliklər həmin qurumların hüquqi statusunu və fəaliyyət çərçivəsini tabe tənzimləmə ilə "
                "uzlaşdırır.",
            ],
            'EN': [
                "The Cabinet of Ministers amended its own decisions in connection with the implementation of "
                "Presidential Decree No. 601 of 20 February 2026 on the reorganisation of the public legal entities "
                "“Media Development Agency of the Republic of Azerbaijan” and “Center for Analysis of International "
                "Relations”.",
                "The amendments align the legal status and scope of activity of those bodies with the subordinate "
                "regulation.",
            ],
        },
                'eqanun': [61859],
    },
    {
        'n': '14',
        'kicker': {'AZ': 'NAZİRLƏR KABİNETİ · SƏRƏNCAMLAR · 18–20 MAY 2026',
                   'EN': 'CABINET · ORDERS · 18–20 MAY 2026'},
        'headline': {'AZ': 'Bir sıra sərəncam son qanunvericiliyin icrasını təmin edir.',
                     'EN': 'A series of orders implements recent legislation.'},
        'body': {
            'AZ': [
                "Kabinet həmçinin bir neçə icra sərəncamı vermişdir: Cinayət Məcəlləsinə və “Terrorçuluğa qarşı "
                "mübarizə haqqında” Qanuna dəyişikliklərin tətbiqi (№ 340s), “Yol hərəkəti haqqında” Qanuna "
                "dəyişikliklər (№ 338s və 331s), İnzibati Xətalar Məcəlləsinə dəyişikliklərin tətbiqi (№ 330s) və "
                "yüksək səviyyəli nümayəndə heyətlərinin səfərləri barədə hesabatlara dair düzəliş (№ 339s).",
                "Ayrı-ayrılıqda texniki xarakter daşısa da, bu sərəncamlar qəbul edilmiş qanun və fərmanların "
                "inzibati praktikaya tətbiqini təmin edir.",
            ],
            'EN': [
                "The Cabinet also issued several implementing orders: the application of amendments to the Criminal "
                "Code and the Law on Combating Terrorism (No. 340s), amendments to the Law on Road Traffic "
                "(Nos. 338s and 331s), the application of amendments to the Code of Administrative Offences (No. 330s), "
                "and a change concerning reports on the visits of high-level delegations (No. 339s).",
                "Individually technical, these orders give effect, in administrative practice, to the laws and decrees "
                "recently adopted.",
            ],
        },
                'eqanun': [61947, 61943, 61944, 61974, 61975],
    },
]

if __name__ == '__main__':
    made = build_issue(BASENAME, CHROME, LEAD, ARTICLES, OUTDIR)
    for f in made:
        print('  wrote', os.path.basename(f))
