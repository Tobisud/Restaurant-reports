import subprocess

def run_script(script_name):
    result = subprocess.run(['python', script_name], capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error running {script_name}: {result.stderr}")
    else:
        print(f"Successfully ran {script_name}")

if __name__ == "__main__":
    # Define the scripts to run
    scripts = ['time_report.py', 'tip_report.py', 'combine_report.py']

    # Run each script in sequence
    for script in scripts:
        run_script(script)
input("Press Enter to exit...")

