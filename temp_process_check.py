from pathlib import Path
from docx import Document
from document_automation_studio.core.controller import ApplicationController
from document_automation_studio.processors.metadata_processor import MetadataValues

root = Path('c:/Users/User/Documents/document_automation_studio')
input_dir = root / 'temp_input'
output_dir = root / 'temp_output'
input_dir.mkdir(exist_ok=True)
out_dir = output_dir
out_dir.mkdir(exist_ok=True)

for name in ['sample.docx']:
    p = input_dir / name
    if not p.exists():
        doc = Document()
        table = doc.add_table(rows=2, cols=2)
        table.rows[0].cells[0].text = 'Effective Date'
        table.rows[0].cells[1].text = '22-07-2026'
        table.rows[1].cells[0].text = 'Prepared By (Name / Designation)'
        table.rows[1].cells[1].text = 'Nur Syafirah / Director'
        doc.save(p)

controller = ApplicationController(config_path=root / 'temp_config.json')
controller.config.processing.input_folder = str(input_dir)
controller.config.processing.output_folder = str(output_dir)
controller.config.processing.process_word_files = True
controller.config.processing.process_excel_files = False
controller.update_processing_settings(controller.config.processing)
controller.set_metadata_values(MetadataValues(effective_date='01-08-2026', prepared_by_name='John Smith', prepared_by_designation='Quality Manager'))
controller.run_batch()
print('done')
