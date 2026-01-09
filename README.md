# Bid Proposal Agent

AI-powered bid proposal analysis tool for civil engineering projects. Streamlines the bidding process by automating document parsing, quantity verification, and proposal review.

## Features

### 1. Proposal Document Parsing
- Upload RFP/bid documents (PDF, Excel)
- AI extracts line items, quantities, requirements
- Identifies project info, deadlines, special requirements

### 2. Bid Proposal Review
- Upload your working bid proposal
- Compare against RFP requirements
- Identify missing items and discrepancies

### 3. Quantity Calculator
- AI-powered quantity extraction from plan sheets
- Supports all civil engineering categories:
  - **Earthwork**: Excavation, Fill, Embankment (CY)
  - **Paving**: HMA, Concrete, Aggregate (TON/SY)
  - **Utilities**: Storm, Sanitary, Water (LF/EA)
  - **Structures**: Retaining walls, Culverts (SF/LF)
  - **Erosion Control**: Silt fence, Seeding (LF/AC)
  - **Traffic Control**: Signs, Markings (EA/LF)
  - **Landscaping**: Trees, Sod, Irrigation (EA/SY)
  - **Survey**: Boundary, Topo, Staking (LS/AC)

### 4. Analysis Reports
- Completeness scoring
- Accuracy verification
- Critical issues and warnings
- Prioritized recommendations
- Export to Word and Excel

## Benefits

- **Faster Bids**: Automated extraction reduces manual reading time
- **More Accurate**: AI quantity verification catches errors
- **More Capacity**: Handle more bids without adding unbillable hours
- **Consistent Process**: Standardized review methodology

## Deployment

### Railway

1. Connect your GitHub repository to Railway
2. Set environment variables:
   - `OPENAI_API_KEY`: Your OpenAI API key
   - `SECRET_KEY`: Flask session secret (generate a random string)
3. Deploy

### Local Development

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows

# Install dependencies
pip install -r requirements.txt

# Set environment variables
export OPENAI_API_KEY=your_key_here
export SECRET_KEY=dev_secret_key

# Run
python app.py
```

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Main web interface |
| `/health` | GET | Health check |
| `/api/parse-proposal` | POST | Parse RFP/bid documents |
| `/api/parse-bid` | POST | Parse working bid proposal |
| `/api/extract-quantities` | POST | Extract quantities from plans |
| `/api/analyze` | POST | Run bid analysis |
| `/api/compare-quantities` | POST | Compare bid vs plan quantities |
| `/api/export/word` | POST | Export Word report |
| `/api/export/excel` | POST | Export Excel quantities |
| `/api/status` | GET | Get session status |
| `/api/clear` | POST | Clear session data |

## Tech Stack

- **Backend**: Flask, Gunicorn
- **AI**: OpenAI GPT-4o (Vision)
- **PDF Processing**: PyMuPDF
- **Excel**: openpyxl
- **Document Export**: python-docx

## Contact

Questions? Contact Jonah Knip - jknip@abonmarche.com
