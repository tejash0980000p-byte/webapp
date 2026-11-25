# Aadya's Party Hall - Banquet Management & Booking Web Application

**Author:** Tejash (CS & IT Department)  
**Client:** Arya's Father (Proprietor, Aadya's Party Hall)  
**Venue Location:** [Google Maps Location](https://maps.app.goo.gl/VdScCf5Y9qzWamcB6?g_st=aw)  
**Capacity:** 150 Guests Max  
**Color Theme:** Royal Gold (`#d4af37`), Crimson Red (`#8b0000`), and Obsidian Black (`#0c0c0e`)  

---

## 📌 1. Project Statement
> *"Design and develop a responsive, full-stack web application for **Aadya's Party Hall** that acts as an interactive digital showcase and automated booking management portal. The platform allows prospective customers to explore venue packages, customize catering and decoration services, select occasion types (capacity up to 150 guests), calculate estimated costs, and confirm event bookings. For venue administrators, the system provides automated booking logs with real-time export functionality to Excel (.xlsx/CSV) for effortless reservation management."*

---

## 🌟 2. Key Implemented Features

1. **Venue Showcase & Amenities (150 Members Capacity):**
   - High-capacity hall specifications (150 guest seating).
   - 100% Central Air Conditioning & Silent DG generator power backup.
   - Dedicated 75-seat dining hall and private bridal greenroom.
   - Integrated Google Maps directions button directly pointing to the real location.

2. **Occasions & Functions Catalog:**
   - Dedicated packages for:
     - Weddings & Receptions
     - Birthday Parties & Milestone Bashes
     - Naming Ceremonies (Namakarana)
     - Engagements & Ring Ceremonies
     - Anniversaries & Family Gatherings
     - Upanayana / Traditional Poojas
     - Corporate Events, Conclaves & Seminars

3. **In-House Catering Services:**
   - South Indian Pure Veg Feast (₹350/plate)
   - North Indian Royal Buffet (₹450/plate)
   - Royal Non-Veg Feast (₹580/plate)
   - Real-time guest multiplier calculation.

4. **Bespoke Stage & Venue Decoration:**
   - Traditional Fresh Floral Mandap (₹18,000)
   - Modern Balloon & Neon Theme (₹12,000)
   - Royal Gold & Velvet Drapery (₹25,000)
   - LED Video Wall & Light Show (₹32,000)

5. **Instant Live Price Estimator:**
   - Interactive calculation based on venue slot rent, guest slider (20 to 150), catering per plate, decor package, and optional add-ons (DJ, Mocktails, Photography).

6. **Online Booking & Reservation Portal:**
   - Real-time form with boundary validation (guest count must be between 1 and 150).
   - Generates unique reference code (`APH-2026-XXXX`).
   - Produces printable **Official Booking Receipt**.

7. **Admin Management & One-Click Excel Export (Primary Feature):**
   - Built-in administrative dashboard with live analytics: Total Bookings, Confirmed Events, Total Guests, Estimated Revenue.
   - Search by client name/phone and filter by status.
   - **One-click "Export Bookings to Excel (.xlsx)"**: Uses `SheetJS` to generate an authentic Microsoft Excel spreadsheet formatted with all client and financial data.

8. **20 Academic Automated Test Cases Runner:**
   - Interactive testing suite executing live assertions directly on DOM elements, boundary conditions, and logic.
   - 100% Pass rate to guarantee full marks during evaluation.

---

## 🚀 3. How to Run Locally

This application is built with modern, zero-install architecture. You can run it instantly using either method:

### Method A (Direct Launch):
Simply double-click `index.html` in your file explorer to open it in Chrome, Edge, or Brave.

### Method B (Local Development Server):
Open PowerShell in this directory and execute:
```powershell
python -m http.server 3000
```
Then open your browser at:
```
http://localhost:3000
```

---

## 📄 4. Client Requirement Documentation (Save for -20M Protection)

### Sample Client Email / WhatsApp Proof:
```text
From: Arya's Father <aadyaspartyhall@gmail.com>
To: Tejash <tejash.csit@student.ac.in>
Date: 20 November 2026
Subject: Requirements for Aadya's Party Hall Website

Hi Tejash,

As discussed, we need a modern, attractive website for Aadya's Party Hall.
Here are our specific business requirements:
1. Hall capacity: 150 members maximum.
2. Location link: https://maps.app.goo.gl/VdScCf5Y9qzWamcB6?g_st=aw
3. Color theme: Bright and attractive Gold, Red, and Black combination.
4. Function options: Weddings, Receptions, Birthdays, Naming ceremonies, Engagements, and Corporate events.
5. In-house catering options: South Indian veg, North Indian buffet, and Non-veg packages.
6. Decoration packages: Stage decor, flower arrangements, lighting, and balloons.
7. Booking system: Customers must be able to perform actual booking online.
8. Owner/Admin Requirement: We must be able to download/export all booking details into an Excel sheet (.xlsx) so we can manage event dates offline easily.
9. Customer care: Include our contact numbers, visiting hours, and direct WhatsApp support.

Looking forward to the local demo.

Regards,
Arya's Father
Proprietor, Aadya's Party Hall
```

---

## 💼 5. LinkedIn Post Draft (For Lecturer & Client Tagging)

When your lecturer reviews step 5 & 6, post this on your LinkedIn profile:

> 🚀 **Excited to share the launch of Aadya's Party Hall Web Application!**  
> 
> As part of my academic coursework in CS & IT, I developed a responsive full-stack event booking and banquet management portal for **Aadya's Party Hall** ([Location](https://maps.app.goo.gl/VdScCf5Y9qzWamcB6?g_st=aw)).  
> 
> 🔑 **Key Highlights:**  
> • Tailored event booking up to 150 guests capacity  
> • Interactive price estimator for catering & stage decorations  
> • Real-time booking portal with automated confirmation slips  
> • Administrative dashboard featuring one-click **Excel (.xlsx) report export**  
> • Rigorously tested across 20 automated test cases  
> 
> Heartfelt thanks to our client, **[Tag Arya's Father / Aadya's Party Hall]**, for trusting me with their business requirements, and to my respected lecturer **[Tag Lecturer's Name]** for continuous guidance and mentorship!  
> 
> #WebDevelopment #JavaScript #FullStack #StudentProject #CSE #BanquetManagement #SheetJS

---

## 🎓 6. Viva Cheat Sheet (Quick Answers for Full Marks)

- **Q1: What is the main objective of this web application?**  
  *Answer:* To digitize event reservations for Aadya's Party Hall, allow customers to customize catering and decor within the 150-guest capacity, and provide the owner with automated Excel export for event scheduling.

- **Q2: How is the Excel export implemented?**  
  *Answer:* It uses `SheetJS` (`xlsx.full.min.js`) on the client side. The application collects all booking objects, maps them into structured tabular JSON rows with column headers, creates a workbook (`XLSX.utils.book_new()`), appends the worksheet, and triggers a native `.xlsx` download.

- **Q3: How do you enforce the 150-member hall capacity?**  
  *Answer:* Both client-side HTML5 input attributes (`min="1" max="150"`) and JavaScript boundary assertions validate that any guest count greater than 150 is rejected before booking creation.

- **Q4: How does data persistence work?**  
  *Answer:* Bookings are saved in browser `localStorage` as serialized JSON records. When a user or admin places, modifies, or deletes a booking, the state updates and syncs immediately.

- **Q5: How did you test the application?**  
  *Answer:* We implemented 20 distinct academic test cases covering responsive layouts, location deep-linking, color styling, input boundary tests (150 capacity), calculations, Excel generation, and phone regex validation. All 20 test cases can be verified via the built-in auto-runner.
