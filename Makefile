.PHONY: all check_python venv install_deps run

# Default target: run the application
all: run

# Check if python is installed; if not, install via Chocolatey
check_python:
	@python --version >nul 2>&1 || (echo Python not found. Installing Python with Chocolatey... && choco install python -y)

# Create virtual environment if it doesn't exist
venv: check_python
	@if not exist "venv\Scripts\python.exe" ( \
		echo Creating virtual environment... && \
		python -m venv venv \
	) else ( \
		echo Virtual environment already exists. \
	)

# Install dependencies from requirements.txt using the venv's pip
install_deps: venv
	@echo Installing Python dependencies...
	@venv\Scripts\pip install -r requirements.txt

# Run main.py using the virtual environment's python interpreter
run: install_deps
	@echo Running main.py...
	@venv\Scripts\python main.py