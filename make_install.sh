pyinstaller src/main.py \
  --paths .venv/lib/python3.9/site-packages:src \
  --add-data "src/presentationMaker/:presentationMaker" \
  --onefile &&
cp dist/main $1 && rm -rf dist