mkdir -p ~/.streamlit/

echo "\
[general]\n\
email = \"contact@3forceconsulting.com\"\n\
" > ~/.streamlit/credentials.toml

echo "\
[server]\n\
headless = true\n\
enableCORS=false\n\
port = $PORT\n\
[theme]\n\
primaryColor = \"#6B7F3B\"\n\
backgroundColor = \"#FFFFFF\"\n\
secondaryBackgroundColor = \"#F5F5F0\"\n\
textColor = \"#262730\"\n\
font = \"sans serif\"\n\
" > ~/.streamlit/config.toml
