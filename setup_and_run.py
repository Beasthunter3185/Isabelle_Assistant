import os
import sys
import json
import subprocess

CONFIG_FILE = "config.json"
REQUIREMENTS_FILE = "requirements.txt"

def load_config() -> dict:
    """Loads configuration settings from config.json."""
    if not os.path.exists(CONFIG_FILE):
        print(f"[ERROR] Config file '{CONFIG_FILE}' not found!")
        sys.exit(1)
    
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)

def check_and_install_requirements():
    """Installs required Python packages if missing."""
    setup_flag = ".setup_complete"
    
    if not os.path.exists(setup_flag):
        print("--- FIRST-TIME SETUP: Installing Python Dependencies ---")
        if os.path.exists(REQUIREMENTS_FILE):
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", REQUIREMENTS_FILE])
                with open(setup_flag, "w") as f:
                    f.write("ready")
                print("--- Python dependencies installed successfully! ---")
            except Exception as e:
                print(f"[ERROR] Failed to install requirements: {e}")
                sys.exit(1)
        else:
            print(f"[WARNING] '{REQUIREMENTS_FILE}' not found. Skipping pip install.")

def ensure_ollama_model(model_name: str):
    """Checks if Ollama is available and pulls the specified model."""
    print(f"--- Checking Ollama setup for model: '{model_name}' ---")
    try:
        # Check if Ollama CLI is accessible
        result = subprocess.run(["ollama", "list"], capture_output=True, text=True)
        if result.returncode == 0:
            if model_name not in result.stdout:
                print(f"Model '{model_name}' not found locally. Pulling now...")
                subprocess.run(["ollama", "pull", model_name], check=True)
                print(f"--- Model '{model_name}' pulled successfully! ---")
            else:
                print(f"--- Model '{model_name}' is ready. ---")
        else:
            print("[WARNING] Could not connect to Ollama service. Make sure Ollama is running.")
    except FileNotFoundError:
        print("[WARNING] Ollama executable not found in system PATH. Please install Ollama.")
    except Exception as e:
        print(f"[WARNING] Error verifying Ollama model: {e}")

def run_server(config: dict):
    """Boots the FastAPI / Uvicorn server using parameters from config.json."""
    import uvicorn
    
    host = config.get("server", {}).get("host", "127.0.0.1")
    port = config.get("server", {}).get("port", 8000)
    
    print(f"--- Starting Isabelle Backend Server on http://{host}:{port} ---")
    uvicorn.run("server:app", host=host, port=port, reload=False)

if __name__ == "__main__":
    # 1. Load dynamic config
    config = load_config()
    
    # 2. Verify virtualenv pip requirements
    check_and_install_requirements()
    
    # 3. Check for the configured Ollama model
    target_model = config.get("ollama", {}).get("model_name", "llama3.1")
    ensure_ollama_model(target_model)
    
    # 4. Start the backend app
    run_server(config)