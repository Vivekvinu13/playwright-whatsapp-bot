# WhatsApp Message Sender + Smart Data Extractor 🤖💬

A Python-based **RPA (Robotic Process Automation)** project that uses **Playwright** to automate WhatsApp Web.

The bot reads contact information from an Excel file, sends personalized WhatsApp messages, captures screenshots of sent messages, extracts the last three incoming messages from each chat, and generates JSON and Excel reports.

## 🎯 Objective

The goal of this project is to demonstrate browser-based RPA using **Playwright with Python**.

The automation interacts with WhatsApp Web similarly to a human user while using Playwright to locate webpage elements, enter messages, send them, and extract information.

## 🔄 Automation Workflow

```text
contacts.xlsx
      ↓
Read contact details
      ↓
Open WhatsApp Web
      ↓
Reuse saved login session
      ↓
Search contact by name or phone
      ↓
Personalize message using {name}
      ↓
Send WhatsApp message
      ↓
Take screenshot of sent message
      ↓
Extract last 3 incoming messages
      ↓
Repeat for next contact
      ↓
Generate JSON + Excel reports
