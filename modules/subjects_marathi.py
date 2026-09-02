# modules/subjects_marathi.py

BA_MAPPING_DATA = [
    (11, 'English', 'इंग्रजी'),
    (12, 'Marathi', 'मराठी'),
    (13, 'Hindi', 'हिंदी'),
    (14, 'Gujarati', 'गुजराती'),
    (15, 'Sanskrit', 'संस्कृत'),
    (16, 'Urdu', 'उर्दू'),
    (18, 'Sindhi', 'सिंधी'),
    (19, 'Arabic', 'अरेबिक'),
    (20, 'French', 'फ्रेंच'),
    (23, 'Pali', 'पाली'),
    (25, 'Persian', 'पर्शियन'),
    (31, 'Economics', 'अर्थशास्त्र'),
    (32, 'History', 'इतिहास'),
    (33, 'Sociology', 'समाजशास्त्र'),
    (34, 'Politics', 'राज्यशास्त्र'),
    (35, 'Philosophy', 'तत्वज्ञान'),
    (36, 'Psychology', 'मानसशास्त्र'),
    (37, 'Geography', 'भूगोल'),
    (38, 'Ancient Indian Culture', 'प्राचीन भारतीय संस्कृती'),
    (39, 'Commerce', 'वाणिज्य'),
    (40, 'Mathematics', 'गणितशास्त्र'),
    (41, 'Statistics', 'संख्याशास्त्र'),
    (42, 'Islamic Studies', 'इस्लामिक अभ्यास'),
    (45, 'Education', 'शिक्षणशास्त्र'),
    (48, 'Rural Development', 'ग्रामीण विकास')
]

BSC_MAPPING_DATA = [
    (40, 'Mathematics', 'गणितशास्त्र'),
    (41, 'Physics', 'भौतिकशास्त्र'),
    (42, 'Chemistry', 'रसायनशास्त्र'),
    (43, 'Statistics', 'संख्याशास्त्र'),
    (44, 'Life Sciences', 'आर्युविज्ञान'),
    (45, 'Botany', 'वनस्पतीशास्त्र'),
    (46, 'Zoology', 'प्राणिशास्त्र'),
    (47, 'Microbiology', 'सूक्ष्मजीवशास्त्र'),
    (48, 'Geology', 'भूशास्त्र'),
    (49, 'Biochemistry', 'जैवरसायन'),
    (50, 'Computer Science', 'संगणक विज्ञान'),
    (51, 'Biotechnology', 'जैवतंत्रज्ञान'),
    (80, 'Mathematics', 'गणितशास्त्र'),
    (81, 'Physics', 'भौतिकशास्त्र'),
    (82, 'Chemistry', 'रसायनशास्त्र'),
    (85, 'Botany', 'वनस्पतीशास्त्र'),
    (86, 'Zoology', 'प्राणिशास्त्र'),
    (87, 'Microbiology', 'सूक्ष्मजीवशास्त्र'),
    (90, 'Computer Science', 'संगणक विज्ञान'),
    (91, 'Biotechnology', 'जैवतंत्रज्ञान')
]

ba_subject_dict = {sub_id: {'name': name, 'namem': namem} for sub_id, name, namem in BA_MAPPING_DATA}
bsc_subject_dict = {sub_id: {'name': name, 'namem': namem} for sub_id, name, namem in BSC_MAPPING_DATA}

# Combined default dictionary
subject_dict = {**ba_subject_dict, **bsc_subject_dict}
