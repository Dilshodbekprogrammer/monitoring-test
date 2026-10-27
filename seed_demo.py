# -*- coding: utf-8 -*-
"""Demo ma'lumotlar: fan, o'qituvchi, savollar, sinf va o'quvchilar yaratadi.

Ishlatish:  python seed_demo.py
O'qituvchi: login=oqituvchi  parol=oqituvchi123
"""
from werkzeug.security import generate_password_hash

import app as A


def main():
    with A.app.app_context():
        if A.q("SELECT 1 FROM users WHERE username='oqituvchi'", one=True):
            print("Demo ma'lumotlar allaqachon mavjud.")
            return
        sid = A.ex("INSERT INTO subjects(name) VALUES('Matematika')")
        tid = A.ex("INSERT INTO users(username,password_hash,role,full_name,created_at) VALUES(?,?,?,?,?)",
                   ("oqituvchi", generate_password_hash("oqituvchi123"), "teacher", "Demo O'qituvchi", A.iso(A.now())))
        A.ex("INSERT INTO teacher_subjects VALUES(?,?)", (tid, sid))
        for i in range(1, 11):
            a, b = i + 1, i + 2
            opts = [str(a + b), str(a + b + 1), str(a + b - 1), str(a + b + 2)]
            A.ex("INSERT INTO questions(teacher_id,subject_id,text,a,b,c,d,correct,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
                 (tid, sid, "%d + %d nechiga teng?" % (a, b), *opts, "A", A.iso(A.now())))
        for n in ["Aliyev Vali", "Karimova Madina", "Tursunov Jasur"]:
            pw = A.make_password()
            A.ex("INSERT INTO users(username,password_hash,pw_plain,role,full_name,class_name,created_at) VALUES(?,?,?,?,?,?,?)",
                 (A.make_username(n), generate_password_hash(pw), pw, "student", n, "7-A", A.iso(A.now())))
        print("Tayyor. O'qituvchi: oqituvchi / oqituvchi123")


if __name__ == "__main__":
    main()
