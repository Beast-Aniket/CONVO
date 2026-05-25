# subjects_marathi.py
mapping_data = [
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

subject_dict = {sub_id: {'name': name, 'namem': namem} for sub_id, name, namem in mapping_data}
