
# Isabelle Assistant 🌸

A cross-platform productivity desktop assistant built with a **Godot 4** transparent overlay frontend, a **Python FastAPI** backend, and local AI integration via **Ollama**. Designed with a clean, aesthetic pastel theme to help manage calendars, tasks, and daily scheduling.

---

## 🛠️ Tech Stack & Architecture

* **Frontend:** Godot Engine 4.x (GDScript, transparent overlay window with click-through mouse passthrough regions).
* **Backend:** Python FastAPI, Uvicorn, MSAL (Microsoft Graph API authentication).
* **AI & Intelligence:** Local LLM routing via Ollama (`llama3.1`).
* **Data Storage:** Local JSON persistence for tasks and configurations.

---

## 📋 Prerequisites

Before running the application, make sure you have the following installed on your system:

1. **Python (3.10 or higher)** installed with `pip`.
2. **Ollama** installed and running locally with your target model:
   ```bash
   ollama run llama3.1
   ```
