from flask import Flask, request, jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
from datetime import datetime
import os
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)
CORS(app)

# PostgreSQL konfiguratsiyasi
app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://admin:ue2F2rR0qrwNm6alG061Gpc7T0jO9dGZ@dpg-d0ib4n0dl3ps738asjbg-a.frankfurt-postgres.render.com/testdb_ko5g'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Ma'lumotlar bazasi modellari
class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    surname = db.Column(db.String(100), nullable=False)
    login = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(50), nullable=False)

class TestFile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    data = db.Column(db.LargeBinary, nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)

class TestResult(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_name = db.Column(db.String(200), nullable=False)
    correct_answers = db.Column(db.Integer, nullable=False)
    total_questions = db.Column(db.Integer, nullable=False)
    date = db.Column(db.DateTime, default=datetime.utcnow)

# API endpointlari
@app.route('/api/verify-student', methods=['POST'])
def verify_student():
    data = request.json
    student = Student.query.filter_by(
        login=data['login'],
        password=data['password']
    ).first()
    
    if student:
        return jsonify({
            'verified': True,
            'name': f"{student.name} {student.surname}"
        })
    return jsonify({'verified': False}), 401

@app.route('/api/students', methods=['GET', 'POST'])
def manage_students():
    if request.method == 'GET':
        students = Student.query.all()
        return jsonify([{
            'name': s.name,
            'surname': s.surname,
            'login': s.login
        } for s in students])
    
    data = request.json
    student = Student(
        name=data['name'],
        surname=data['surname'],
        login=data['login'],
        password=data['password']
    )
    db.session.add(student)
    db.session.commit()
    return jsonify({'message': 'Student added successfully'})

@app.route('/api/students/<login>', methods=['DELETE'])
def delete_student(login):
    student = Student.query.filter_by(login=login).first()
    if student:
        db.session.delete(student)
        db.session.commit()
        return jsonify({'message': 'Student deleted successfully'})
    return jsonify({'message': 'Student not found'}), 404

@app.route('/api/upload-test', methods=['POST'])
def upload_test():
    if 'file' not in request.files:
        return jsonify({'message': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'message': 'No file selected'}), 400
    
    filename = secure_filename(file.filename)
    file_data = file.read()
    
    test_file = TestFile(filename=filename, data=file_data)
    db.session.add(test_file)
    db.session.commit()
    
    return jsonify({'message': 'Test file uploaded successfully'})

@app.route('/api/get-test/<int:test_id>', methods=['GET'])
def get_test(test_id):
    test_file = TestFile.query.get_or_404(test_id)
    
    # Vaqtinchalik fayl yaratish
    temp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    temp.write(test_file.data)
    temp.close()
    
    try:
        with open(temp.name, 'rb') as f:
            data = f.read()
        os.unlink(temp.name)
        return data, 200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': f'attachment; filename={test_file.filename}'
        }
    except Exception as e:
        if os.path.exists(temp.name):
            os.unlink(temp.name)
        return jsonify({'message': str(e)}), 500

@app.route('/api/results', methods=['GET', 'POST'])
def manage_results():
    if request.method == 'GET':
        results = TestResult.query.order_by(TestResult.date.desc()).all()
        return jsonify([{
            'student_name': r.student_name,
            'correct_answers': r.correct_answers,
            'total_questions': r.total_questions,
            'date': r.date.strftime('%Y-%m-%d %H:%M:%S')
        } for r in results])
    
    data = request.json
    result = TestResult(
        student_name=data['student_name'],
        correct_answers=data['correct_answers'],
        total_questions=data['total_questions'],
        date=datetime.strptime(data['date'], '%Y-%m-%d %H:%M:%S')
    )
    db.session.add(result)
    db.session.commit()
    return jsonify({'message': 'Result saved successfully'})

@app.route('/api/results/export', methods=['GET'])
def export_results():
    import openpyxl
    from io import BytesIO
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Test Results"
    
    # Ustunlar
    headers = ['Student', 'Correct Answers', 'Total Questions', 'Percentage', 'Date']
    ws.append(headers)
    
    # Ma'lumotlar
    results = TestResult.query.order_by(TestResult.date.desc()).all()
    for r in results:
        percentage = round((r.correct_answers / r.total_questions) * 100, 2)
        ws.append([
            r.student_name,
            r.correct_answers,
            r.total_questions,
            f"{percentage}%",
            r.date.strftime('%Y-%m-%d %H:%M:%S')
        ])
    
    # Fayl yaratish
    excel_file = BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    
    return excel_file.getvalue(), 200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename=test_results.xlsx'
    }

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port) 