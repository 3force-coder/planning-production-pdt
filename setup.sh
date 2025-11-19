mkdir -p ~/.streamlit/

echo "\
[general]\n\
email = \"votre@email.com\"\n\
" > ~/.streamlit/credentials.toml

echo "\
[server]\n\
headless = true\n\
enableCORS=false\n\
port = $PORT\n\
[theme]\n\
primaryColor = \"#2E7D32\"\n\
backgroundColor = \"#FFFFFF\"\n\
secondaryBackgroundColor = \"#F0F2F6\"\n\
textColor = \"#262730\"\n\
" > ~/.streamlit/config.toml
