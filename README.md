# 🏫 Teacher Allocation System

> A Google Apps Script automation tool to streamline teacher-to-course allocation based on skills, category, and real-time availability.

## 🌟 Features

* **Smart Filtering:** Filter courses by category (e.g., Roblox, Python) to find the right curriculum.
* **Availability Tracking:** Automatically checks teacher schedules to prevent double-booking.
* **Skill Matching:** Matches teachers based on specific expertise (Expert/Beginner) using a "Category + Exception" logic.
* **Visual Dashboard:** Color-coded status indicators (✅ Available, ⚠️ Not Available) for quick decision-making.
* **One-Click Assignment:** Assigns teachers and updates their schedule instantly.

## 🛠️ How It Works

This system is built on **Google Sheets** and **Google Apps Script**. It uses a logic-first approach:

1.  **Select a Course:** The user filters and selects a course from the dashboard.
2.  **Search:** The script scans the `Teacher_Category_Master` to find teachers qualified for that category.
3.  **Check Schedule:** It cross-references the qualified teachers against their personal Availability Sheets.
4.  **Result:** Displays a ranked list of available teachers with their workload and skill level.

## 📂 File Structure

* `Code.gs`: Contains all the backend logic, including search algorithms, availability checks, and UI triggers.
* **Google Sheets Data Sources:**
    * `Dashboard`: The main UI for users.
    * `Courses_Master`: Database of all 450+ courses.
    * `Teacher_Category_Master`: Links teachers to broad categories (e.g., Varsha -> Roblox).
    * `Assignments_Log`: Tracks history of all assignments.

## 🚀 Setup Guide

1.  **Copy the Sheet:** Create a new Google Sheet.
2.  **Open Script Editor:** Go to `Extensions` > `Apps Script`.
3.  **Paste Code:** Copy the content of `Code.gs` from this repository into the script editor.
4.  **Create Sheets:** Ensure you have the required tabs (`Dashboard`, `Courses_Master`, etc.) named exactly as in the code.
5.  **Run Setup:** Run the `onOpen` function to create the custom menu.

## 📸 Screenshots

<img width="1608" height="677" alt="image" src="https://github.com/user-attachments/assets/559fd0e3-2ec8-4129-ad9d-37f21679533a" />


## 📄 License

This project is for internal use at Create & Learn.
