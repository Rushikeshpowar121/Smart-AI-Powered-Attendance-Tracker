# Attendance Bot

This project is a voice-enabled Attendance Bot designed to take attendance for predefined subjects. It uses Python libraries for speech recognition, text-to-speech conversion, and file handling. The bot saves attendance data in an Excel file for future reference.

---

## Features
- **Voice Greetings:** Welcomes the user based on the time of day.
- **Subject Selection:** Lets the user select a subject from a predefined list.
- **Speech Recognition:** Takes attendance through voice commands ("present" or "absent").
- **Attendance Logging:** Logs absentee information in an Excel file organized by subject.
- **Excel File Handling:** Appends data to an existing sheet or creates a new sheet if the subject is not present.
- **Interactive Prompts:** Allows the user to continue or exit the program.

---

## Prerequisites
### **Libraries Used**
- `time` - For time and date handling.
- `speech_recognition` - For recognizing voice commands.
- `pyttsx3` - For text-to-speech conversion.
- `datetime` - To fetch the current time and date.
- `pandas` - For handling Excel data.
- `openpyxl` - For working with Excel files.
- `os` and `sys` - For system-level operations.

### **Installation of Required Libraries**
Install the required libraries using pip:
```bash
pip install speechrecognition pyttsx3 pandas openpyxl
```

---

## How It Works

### **1. `hello()` Function**
This function greets the user based on the time of day:
- Morning: "Good morning, sir."
- Afternoon: "Good afternoon, sir."
- Evening: "Good evening, sir."

### **2. `thank_you()` Function**
Exits the program gracefully.

### **3. `choose_subject()` Function**
Prompts the user to choose a subject for attendance:
- Lists predefined subjects such as "Management," "Programming with Python," and others.
- Allows the user to exit the program by selecting "6. Exit."
- Returns the name of the chosen subject.

### **4. `recognize_speech()` Function**
Listens to the user's voice input and converts it to text using the `speech_recognition` library. If no input is detected, it simply returns `None`.

### **5. `take_attendance(subject)` Function**
- Takes attendance for roll numbers 1 to 4 using voice input.
- Recognizes "present" as a response; otherwise, marks the roll number as absent.
- Returns a list of absent roll numbers.

### **6. `pronounce_absent(absent_numbers, subject_name)` Function**
- Logs absentee information in an Excel file (`test.xlsx`):
  - Checks if the file exists.
  - Creates a new sheet for the subject if it doesn’t exist or appends data to an existing sheet.
- Reads the list of absentees aloud using the `pyttsx3` library.
- Asks the user if they want to continue or exit the program.

### **7. `speak(text)` Function**
- Converts the provided text to speech using the `pyttsx3` library.

---

## File Handling
- **Excel File (`test.xlsx`):**
  - The program stores absentee data in an Excel file with columns `Date`, `Time`, and `Absentees`.
  - Each subject has its own sheet.
  - Data is appended if the sheet already exists.

---

## Execution
1. **Run the Program:**
   ```bash
   python attendance_bot.py
   ```
2. **Interaction Flow:**
   - The bot greets you.
   - Prompts you to choose a subject.
   - Takes attendance for roll numbers through voice input.
   - Logs absentee data in an Excel file.
   - Asks whether to continue or exit.

---

## Example Usage
### **Input:**
1. Select "1. Management" as the subject.
2. Respond with "present" or stay silent for each roll number prompt.

### **Output:**
- Voice response for attendance status.
- ![image](https://github.com/user-attachments/assets/8ddd983e-1865-4f59-9a33-c86c8f60aae3)
- ![image](https://github.com/user-attachments/assets/c7138dc9-ff4f-4885-ac0a-1b908489b5b4)


- Absentee list logged in `test.xlsx` under the "Management" sheet.
- ![image](https://github.com/user-attachments/assets/55ce411f-c9e8-46d2-af81-975191b1f6b8)


---

## Limitations
- Requires a functional microphone for speech input.
- Works with predefined subjects only.
- Limited to recognizing "present" as the attendance marker.
- Attendance range is currently hardcoded (roll numbers 1 to 4).

---

## Future Improvements
- Add dynamic subject handling.
- Expand roll number range.
- Improve speech recognition accuracy.
- Add a GUI for easier interaction.
- Enhance error handling for edge cases.

---
