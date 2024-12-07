# Word to Excel Converter with SQL Schema Generator

A Python utility that converts Word documents to Excel and automatically generates SQL schema definitions based on the data structure.

## Features

- 📄 Convert Word documents (.docx) to Excel (.xlsx)
- 📊 Preserve formatting and table structures
- 🔍 Intelligent data type detection
- 📝 Generate SQL schemas for multiple database dialects
- 📊 Detailed column analysis and metadata
- ✨ Support for PostgreSQL, MySQL, SQLite, and MSSQL

## Prerequisites

```bash
pip install -r requirements.txt
```

Required Python packages:
- python-docx
- openpyxl
- pandas
- typing

## Usage

1. Update the file paths in `convert.py`:
```python
word_file_path = "/path/to/your/document.docx"
excel_file_path = "output/WordToExcel.xlsx"
```

2. Run the script:
```bash
python convert.py
```

3. The script will generate:
   - Excel file with the converted Word document
   - SQL schema files for PostgreSQL and MySQL
   - Detailed column analysis in JSON format

## Output Files

- `output/WordToExcel.xlsx`: Converted Excel file
- `output/schema_postgresql.sql`: PostgreSQL schema
- `output/schema_mysql.sql`: MySQL schema
- `output/column_analysis.json`: Detailed column metadata

## SQL Type Mapping

The script intelligently maps data types based on content analysis:

### Integer Types
- `SMALLINT`: Values < 32,767
- `INTEGER/INT`: Values < 2,147,483,647
- `BIGINT`: Larger values

### Text Types
- PostgreSQL:
  - `VARCHAR(n)`: Text up to 255 characters
  - `TEXT`: Longer text
- MySQL:
  - `VARCHAR(n)`: Text up to 65,535 characters
  - `MEDIUMTEXT`: Text up to 16MB
  - `LONGTEXT`: Text up to 4GB

### Date/Time Types
- PostgreSQL: `TIMESTAMP`
- MySQL: `DATETIME`

### Decimal Types
- PostgreSQL: `NUMERIC(precision,scale)`
- Others: `DECIMAL(precision,scale)`

## Column Analysis

The JSON analysis includes:
- Original column names
- Null value statistics
- Data ranges (min/max values)
- Text length statistics
- Data quality metrics

## Example Output

```sql
CREATE TABLE example_table (
    id INTEGER NOT NULL,
    name VARCHAR(100) NULL,
    created_at TIMESTAMP NULL,
    amount NUMERIC(18,2) NULL
);
```

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is licensed under the MIT License - see the LICENSE file for details.