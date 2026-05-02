# chatbot-telegram
# Telegram Chatbot – AI-powered Customer Support
## Overview
This project is a Telegram chatbot designed to automate customer support using basic AI techniques, mainly Natural Language Processing (NLP).
Instead of relying on fixed commands, the bot tries to understand what users actually mean, even when the input is not perfect.
The goal is to simulate a more natural interaction while still keeping the system lightweight and practical.

## AI / NLP Approach
The core of this project is a simple but effective NLP pipeline.
### Text normalization
User input is cleaned before processing:
  - Remove extra spaces
  - Standardize casing
  - andle small typos

-> This step improves consistency and helps the bot understand messy real-world input.
### Intent detection (rule-based NLP)
Instead of using heavy models, the bot uses:
  - Keyword matching
  - Category mapping
  - Custom logic

-> This approach is lightweight but still effective for structured customer queries.
### Context handling
The bot keeps track of previous messages to handle follow-up questions.
#### For example:
  “Show me product A”    
  “What about type B?”
  
-> It understands that the second question is related to the first one, which creates a more natural conversation flow.
### Information extraction
From a single user message, the bot can extract:
  - Product name or code
  - Pricing intent
  - Promotion-related queries

-> This allows flexible queries instead of forcing users to follow strict commands.

## Key Features
### Product search & pricing
  - Extract product information from text
  - Match with dataset
  - Return price in a clear format
### Promotion handling
  - Search promotions by product
  - Translate conditions into simple language
  - Check eligibility based on logic rules
### Smart response generation
Instead of returning raw data, the bot:
  - Filters unnecessary information
  - Highlights key details
  - Keeps responses concise
### Email extraction & logging
  - Detect email from user messages
  - Send data via SMTP
  - Store conversation logs for tracking

## Data Handling
The system uses Excel as a lightweight data source:
  - products.xlsx
  - promotions.xlsx

-> This design keeps the system simple, flexible, and easy to update without modifying code.

## Technologies
  - Python
  - Telegram Bot API
  - Pandas
  - smtplib
  - Rule-based NLP
## Project Structure
  chatbot-telegram/
   - │── app.py
   - │── products.xlsx
   - │── promotions.xlsx
   - │── README.md

## Future Improvements
  - Upgrade NLP to LLM-based understanding
  - Replace Excel with database or external API
  - Add recommendation system
  - Deploy to cloud for real usage
