import tkinter as tk
from tkinter import scrolledtext, filedialog
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button, Entry
import docx2txt
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import pymongo

# Function to connect to MongoDB and insert document
def connect_to_mongodb_and_insert(document):
    client = pymongo.MongoClient("mongodb://localhost:27017/")
    db = client["Final_Resume_Scanner"]
    collection = db["resume_scan"]
    collection.insert_one(document)
    client.close()

# Function to process the similarity calculation and display results
def process_similarity():
    job_description = docx2txt.process(job_desc_path)
    resume = docx2txt.process(resume_path)

    content = [job_description, resume]
    cv = CountVectorizer()
    matrix = cv.fit_transform(content)
    similarity_matrix = cosine_similarity(matrix)

    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, "\n\nResume Content:\n", 'heading')
    result_text.insert(tk.END, resume, 'content')

    matching_percentage = similarity_matrix[1][0] * 100
    percentage_box.delete(1.0, tk.END)
    percentage_box.insert(tk.END, 'Resume matches by: {:.2f}%'.format(matching_percentage))
    percentage_box.tag_configure('center', justify='center', font=("Arial", 16, "bold"))  # Increase font size
    percentage_box.tag_add('center', '1.0', 'end')  # Apply center alignment to the whole text

    resume_info = {
        "job_description": job_description,
        "resume_text": resume,
        "similarity_score": matching_percentage
    }

    connect_to_mongodb_and_insert(resume_info)

# Function to select the job description file
def select_job_description():
    global job_desc_path
    job_desc_path = filedialog.askopenfilename(filetypes=[("Document Files", "*.docx")])
    job_desc_entry.delete(0, tk.END)
    job_desc_entry.insert(0, job_desc_path)

# Function to select the resume file
def select_resume():
    global resume_path
    resume_path = filedialog.askopenfilename(filetypes=[("Document Files", "*.docx")])
    resume_entry.delete(0, tk.END)
    resume_entry.insert(0, resume_path)

def on_enter(e):
    e.widget.config(bootstyle='info-outline')  # Change style on mouse enter

def on_leave(e):
    e.widget.config(bootstyle='secondary-outline')  # Change style on mouse leave

# Initialize main window with ttkbootstrap style
style = Style(theme='vapor')
root = style.master
root.title("Resume Similarity Checker")
root.geometry("800x600")  # Set initial window size

# Create header frame
header_frame = tk.Frame(root, bg="#023458", padx=20, pady=10)
header_frame.pack(fill=tk.X)

# Add logo image to the header
try:
    logo = tk.PhotoImage(file="profile (1).png")  # Replace with your logo path
    logo_label = tk.Label(header_frame, image=logo, bg="#023458")
    logo_label.image = logo  # Keep a reference to the image object
    logo_label.pack(side=tk.LEFT, padx=10)
except Exception as e:
    print(f"Error loading image: {e}")

# Header text
header_label = tk.Label(header_frame, text="RESUME CHECKER", bg="#023458", fg="white", font=("Comic Sans MS", 28, "bold"))
header_label.pack(side=tk.LEFT)

# Create frame for file upload buttons and inputs
upload_frame = tk.Frame(root, bg="#333", padx=20, pady=20)
upload_frame.pack(padx=20, pady=10, fill=tk.X)

# Job description upload button
job_desc_button = Button(upload_frame, text="Select Job Description", command=select_job_description, bootstyle='secondary-outline', width=25, padding=(10, 10))  # Increased font size
job_desc_button.grid(row=0, column=0, pady=10, sticky='ew')
job_desc_button.bind("<Enter>", on_enter)
job_desc_button.bind("<Leave>", on_leave)

# Job description entry
job_desc_entry = Entry(upload_frame, font=("Arial", 14), width=80)
job_desc_entry.grid(row=0, column=1, padx=(10, 40), pady=10, sticky='ew')

# Resume upload button
resume_button = Button(upload_frame, text="Select Resume", command=select_resume, bootstyle='secondary-outline', width=25, padding=(10, 10))  # Increased font size
resume_button.grid(row=1, column=0, pady=10, sticky='ew')
resume_button.bind("<Enter>", on_enter)
resume_button.bind("<Leave>", on_leave)

# Resume entry
resume_entry = Entry(upload_frame, font=("Arial", 14), width=80)
resume_entry.grid(row=1, column=1, padx=(10, 40), pady=10, sticky='ew')

# Create frame for displaying results
result_frame = tk.Frame(root, bg="#222", borderwidth=2, relief="groove", padx=20, pady=20)
result_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(10, 10))

result_text = scrolledtext.ScrolledText(result_frame, width=60, height=20, wrap=tk.WORD)
result_text.pack(pady=10, expand=True, fill=tk.BOTH)  # Make the text box expand to fill available space
result_text.tag_config('heading', justify='center', font=("Arial", 14, "bold"))
result_text.tag_config('content', font=("Arial", 12))

# Create frame for displaying matching percentage
percentage_frame = tk.Frame(root, bg="#333")
percentage_frame.pack(padx=20, pady=(0, 20))

percentage_box = scrolledtext.ScrolledText(percentage_frame, width=40, height=4, wrap=tk.WORD)
percentage_box.pack(pady=5)
percentage_box.tag_configure('center', justify='center', font=("Arial", 16, "bold"))  # Increase font size
percentage_box.tag_add('center', '1.0', 'end')  # Apply center alignment to the whole text

# Calculate similarity button
calculate_button = Button(root, text="Calculate Similarity", command=process_similarity, bootstyle='secondary-outline', width=25, padding=(10, 10))  # Increased font size
calculate_button.pack(pady=(0, 20))
calculate_button.bind("<Enter>", on_enter)
calculate_button.bind("<Leave>", on_leave)

root.mainloop()
