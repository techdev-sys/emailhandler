# RBZ BSD Automation Bot

An intelligent Python-based automation system for the Reserve Bank of Zimbabwe (RBZ) Banking Supervision Department (BSD). This bot automates the process of retrieving, validating, and routing regulatory banking returns from Outlook emails to SharePoint.

## features

- **Intelligent Routing**: Automatically classifies and routes returns (BSD 2, 3, and 4) to their respective destinations.
- **Strict Validation**: Performs deep structural analysis of Excel files to ensure compliance with regulatory formats (e.g., verifying composite BSD 2/3 reports).
- **Keyword Analysis**: Employs a score-based keyword system to accurately identify BSD 4 returns.
- **Outlook Integration**: Supports both standard Microsoft Outlook COM and advanced **Redemption** (RDOSession) for high-speed, stable email processing.
- **Recursive Scan**: Capable of scanning attachments within forwarded messages and embedded items.
- **Historical Scanning**: Includes a specialized mode for re-processing or auditing historical returns for specific dates.
- **Automatic Logging**: maintains a detailed `processed_log.txt` and real-time audit logs of all classification decisions.

## Installation

### Prerequisites
- Python 3.8 or higher.
- Microsoft Outlook installed and configured.
- (Optional but Recommended) [Redemption](http://www.dimastr.com/redemption/home.htm) for improved stability and speed.

### Setup
1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd RBZ_Auto_Bot
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure paths:
   Edit `config.py` to set your `BASE_SHAREPOINT_PATH` (the local sync path of your SharePoint directory).

## Usage

Run the main script to start the bot:

```bash
python main_bot.py
```

### Modes of Operation
1. **Live Monitor**: Continuously scans your inbox for incoming returns every few seconds.
2. **Historical Scan**: Allows you to pick an account and a specific date to scan for missing or late submissions.

## Project Structure
- `main_bot.py`: The core engine for email processing and routing.
- `excel_validator.py`: Contains the logic for structural and keyword-based Excel validation.
- `bank_rules.py`: Defines rules for mapping emails/files to specific financial institutions.
- `config.py`: Global configuration and path settings.
- `.gitignore`: Configured to keep the repository clean of logs, environments, and sensitive data.

## License
MIT
