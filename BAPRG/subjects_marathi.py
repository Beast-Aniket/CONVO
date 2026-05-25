# subjects_marathi.py
mapping_data = [
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
subject_dict = {sub_id: {'name': name, 'namem': namem} for sub_id, name, namem in mapping_data}
