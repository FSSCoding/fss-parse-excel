#!/usr/bin/env python3
"""
Excel Toolkit Installation Script
Professional installation with dependency management and system integration
"""

import subprocess
import sys
import os
import shutil
from pathlib import Path

def run_command(command, description="Running command"):
    """Run a shell command with error handling."""
    print(f"🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"❌ Error: {result.stderr}")
            return False
        if result.stdout.strip():
            print(f"✅ {result.stdout.strip()}")
        return True
    except Exception as e:
        print(f"❌ Exception: {e}")
        return False

def check_python_version():
    """Check if Python version is compatible."""
    if sys.version_info < (3, 8):
        print("❌ Python 3.8 or higher is required")
        print(f"   Current version: {sys.version}")
        return False
    print(f"✅ Python version: {sys.version}")
    return True

def install_dependencies():
    """Install required Python packages."""
    print("📦 Installing dependencies...")
    
    # Core dependencies
    dependencies = [
        "openpyxl>=3.1.0",
        "pandas>=1.5.0", 
        "xlrd>=2.0.1",
        "PyYAML>=6.0",
        "click>=8.0.0",
        "rich>=12.0.0",
        "tabulate>=0.9.0"
    ]
    
    for dep in dependencies:
        if not run_command(f"pip install '{dep}'", f"Installing {dep.split('>=')[0]}"):
            return False
    
    return True

def create_global_command():
    """Create global 'excel' command."""
    print("🔗 Creating global command...")
    
    # Get the current script directory
    excel_dir = Path(__file__).parent.resolve()
    
    # Create executable script
    excel_script = excel_dir / "bin" / "excel"
    excel_script.parent.mkdir(exist_ok=True)
    
    script_content = f'''#!/usr/bin/env python3
"""
Excel Toolkit Global Command
"""
import sys
import os

# Add the src directory to Python path
excel_dir = r"{excel_dir}"
sys.path.insert(0, os.path.join(excel_dir, "src"))

# Import and run the main module
try:
    from excel_engine import main
    if __name__ == "__main__":
        main()
except ImportError as e:
    print(f"❌ Error: Cannot import excel_engine: {{e}}")
    print(f"   Excel directory: {{excel_dir}}")
    print(f"   Python path: {{sys.path}}")
    sys.exit(1)
'''
    
    with open(excel_script, 'w') as f:
        f.write(script_content)
    
    # Make executable
    os.chmod(excel_script, 0o755)
    
    # Try to add to system PATH
    try:
        # Check if ~/bin exists and is in PATH
        home_bin = Path.home() / "bin"
        home_bin.mkdir(exist_ok=True)
        
        # Create symlink
        global_excel = home_bin / "excel"
        if global_excel.exists():
            global_excel.unlink()
        global_excel.symlink_to(excel_script)
        
        print(f"✅ Global command created: {global_excel}")
        print(f"   Make sure {home_bin} is in your PATH")
        
        return True
    except Exception as e:
        print(f"⚠️  Could not create global command: {e}")
        print(f"   You can run directly: {excel_script}")
        return False

def verify_installation():
    """Verify the installation works."""
    print("🧪 Verifying installation...")
    
    # Test imports
    test_imports = [
        ("openpyxl", "Excel .xlsx support"),
        ("pandas", "Data processing"),
        ("yaml", "Configuration files"),
        ("click", "CLI framework"),
        ("rich", "Rich terminal output")
    ]
    
    for module, description in test_imports:
        try:
            __import__(module)
            print(f"✅ {description}: {module}")
        except ImportError:
            print(f"❌ {description}: {module} not available")
            return False
    
    # Test our modules
    excel_dir = Path(__file__).parent.resolve()
    sys.path.insert(0, str(excel_dir / "src"))
    
    try:
        import excel_engine
        print("✅ Excel engine module loaded")
    except ImportError as e:
        print(f"❌ Excel engine module failed: {e}")
        return False
    
    return True

def main():
    """Main installation process."""
    print("🚀 Excel Toolkit Installation")
    print("=" * 50)
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Install dependencies
    if not install_dependencies():
        print("❌ Dependency installation failed")
        sys.exit(1)
    
    # Create global command
    create_global_command()
    
    # Verify installation
    if not verify_installation():
        print("❌ Installation verification failed")
        sys.exit(1)
    
    print("=" * 50)
    print("✅ Excel Toolkit installation complete!")
    print("")
    print("📋 Quick Start:")
    print("   excel convert data.xlsx data.csv")
    print("   excel edit data.xlsx --cell A1 'New Value'")
    print("   excel query data.xlsx --sheet Sales")
    print("")
    print("📚 Documentation: README.md")
    print("🧪 Run tests: python -m pytest tests/")

if __name__ == "__main__":
    main()