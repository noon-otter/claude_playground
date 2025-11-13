# Python Version Requirements

## ⚠️ REQUIRED: Python 3.11 or 3.12 Only

This project **requires** Python 3.11 or 3.12 for stability.

- ✅ **Python 3.11** (most stable)
- ✅ **Python 3.12** (recommended)
- ❌ **Python 3.13+** (not supported - will be rejected by dev.sh)
- ❌ **Python 3.10 and below** (not supported)

## Your Current Python Version

Check your Python version:
```bash
python3 --version
```

## Using a Stable Python Version

### Option 1: Install Python 3.12 (Recommended)

**On macOS:**
```bash
brew install python@3.12
```

**On Ubuntu/Debian:**
```bash
sudo apt install python3.12 python3.12-venv
```

### Option 2: Create venv with Specific Python Version

If you have multiple Python versions installed:

```bash
# Remove existing venv if it exists
rm -rf venv

# Create venv with Python 3.12
python3.12 -m venv venv

# Now run dev.sh
./dev.sh
```

### Option 3: Use pyenv (Advanced)

```bash
# Install pyenv
brew install pyenv  # macOS
# OR: curl https://pyenv.run | bash  # Linux

# Install Python 3.12
pyenv install 3.12.0

# Set as local version for this project
cd excel_addin
pyenv local 3.12.0

# Create venv
python -m venv venv

# Run dev script
./dev.sh
```

## Troubleshooting

### "pydantic-core build failed"
This means your Python version is too new. Use Python 3.11 or 3.12.

### "psycopg build failed"
Make sure you're using the latest requirements.txt from this repo.

### Still having issues?
Delete your venv and recreate with a stable Python version:
```bash
rm -rf venv
python3.12 -m venv venv
./dev.sh
```
