from setuptools import setup, find_packages
import os
import sys  # Missing import

# Read long description from README
with open('README.md', 'r', encoding='utf-8') as f:
    long_description = f.read()

# Get dependencies
with open('requirements.txt', 'r') as f:
    requirements = f.read().splitlines()

# Common setup configuration
setup_params = {
    "name": "ReportSender",
    "version": "1.0.0",
    "author": "eliki",
    "author_email": "elijahkiragum@gmail.com",
    "description": "Automated student report sender to parents",
    "long_description": long_description,
    "long_description_content_type": "text/markdown",
    "packages": find_packages(),
    "include_package_data": True,
    "install_requires": requirements,
    "entry_points": {
        'gui_scripts': [
            'reportsender=report_sender.main:main',
        ],
    },
    "classifiers": [
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Operating System :: Microsoft :: Windows",
        "Operating System :: POSIX :: Linux",
        "Operating System :: MacOS",
    ],
    "python_requires": '>=3.6',
    "keywords": 'education reports email automation',
}

# Windows-specific setup
if sys.platform == 'win32':
    try:
        import winshell
        from win32com.client import Dispatch
        
        # Modify setup parameters
        setup_params.update({
            "options": {
                "build_exe": {
                    "include_msvcr": True,  # Include VC++ runtime
                }
            }
        })
        
        # Post-install script (will run after setup completes)
        def create_shortcut():
            desktop = winshell.desktop()
            path = os.path.join(desktop, "Report Sender.lnk")
            target = os.path.join(sys.prefix, "Scripts", "reportsender.exe")
            
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(path)
            shortcut.Targetpath = target
            shortcut.WorkingDirectory = os.path.dirname(target)
            shortcut.IconLocation = target  # Use executable's icon
            shortcut.save()
        
        setup_params.update({"cmdclass": {
            "install": lambda: (setup_params.get("cmdclass", {}).get("install", lambda: None)(), create_shortcut())
        }})
        
    except ImportError:
        print("Note: Windows-specific dependencies not found. Shortcut won't be created.")

setup(**setup_params)