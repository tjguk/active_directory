IF NOT EXIST _static mkdir _static
IF NOT EXIST _static mkdir _templates
python -c "import sphinx; sphinx.main ([None, '-b', 'html', '.', '.\_build'])"
pause