import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.decomposition import NMF
import re
from collections import defaultdict

# Urgency classification keywords and rules
URGENCY_RULES = {
    'RED_ALERT': {
        'keywords': [
            'cfpb', 'attorney general', 'ag office', 'class action', 'lawsuit',
            'legal action', 'arbitration', 'legal counsel', 'legal demand',
            'regulatory agency', 'licensing body', 'federal agency'
        ],
        'timeframe': '24hrs',
        'description': 'Imminent or severe legal action; significant financial/liability exposure'
    },
    'HIGH': {
        'keywords': [
            'news', 'media', 'lawyer', 'attorney', 'bbb', 'better business bureau',
            'illinois shines', 'il shines', 'executive leadership', 'personal safety',
            'property damage', 'harm'
        ],
        'timeframe': '24hrs',
        'description': 'Identified risks or potential legal action'
    },
    'MEDIUM': {
        'keywords': [
            'urgent', 'asap', 'immediate', 'escalate', 'supervisor',
            'manager', 'complaint'
        ],
        'timeframe': '48hrs',
        'description': 'Standard escalation'
    }
}

def preprocess(text):
    """Enhanced preprocessing with urgency keyword preservation"""
    if pd.isna(text):
        return ""
    text = str(text).lower()
    
    # Preserve multi-word terms
    for urgency_level in URGENCY_RULES.values():
        for keyword in urgency_level['keywords']:
            if keyword in text:
                text = text.replace(keyword, keyword.replace(' ', '_'))
    
    # Basic cleaning
    text = re.sub(r'[^a-zA-Z0-9\s_]', ' ', text)
    return ' '.join(text.split()).strip()

def determine_urgency(text):
    """Determine urgency level based on keywords and context"""
    # Handle null values and non-string types
    if pd.isna(text):
        return 'NORMAL', 'No comment provided'
    text = text.lower()
    
    # Check for Red Alert conditions
    for keyword in URGENCY_RULES['RED_ALERT']['keywords']:
        if keyword in text:
            return 'RED_ALERT', URGENCY_RULES['RED_ALERT']['description']
    
    # Check for High Alert conditions
    for keyword in URGENCY_RULES['HIGH']['keywords']:
        if keyword in text:
            return 'HIGH', URGENCY_RULES['HIGH']['description']
    
    # Default to Medium if any urgency keywords found
    for keyword in URGENCY_RULES['MEDIUM']['keywords']:
        if keyword in text:
            return 'MEDIUM', URGENCY_RULES['MEDIUM']['description']
    
    return 'NORMAL', 'Standard ticket'

def generate_topic_name(keywords):
    """Generate a descriptive name for a topic based on its keywords"""
    issue_terms = {
        'inverter': 'Inverter Issues',
        'panel': 'Panel Problems',
        'roof': 'Roof-related Issues',
        'bill': 'Billing Concerns',
        'app': 'App/Portal Issues',
        'install': 'Installation Issues',
        'power': 'Power Generation',
        'produce': 'Production Issues',
        'payment': 'Payment Issues',
        'refund': 'Refund Requests',
        'legal': 'Legal Issues',
        'safety': 'Safety Concerns'
    }
    
    for term, category in issue_terms.items():
        if term in ' '.join(keywords):
            return category
    
    return f"{keywords[0].title()} {keywords[1].title()} Issues"

# Load and process data
df = pd.read_excel("comments.xlsx", engine='openpyxl')
# Convert comments column to string type and handle NaN values
df['comments'] = df['comments'].astype(str)
processed_comments = [preprocess(comment) for comment in df['comments']]

# Create TF-IDF representation
vectorizer = TfidfVectorizer(
    max_features=1000,
    stop_words='english',
    ngram_range=(1, 2),
    max_df=0.9,
    min_df=2
)
tfidf = vectorizer.fit_transform(processed_comments)
feature_names = vectorizer.get_feature_names_out()

# Apply topic modeling
n_topics = 15
nmf = NMF(n_components=n_topics, random_state=42)
nmf_output = nmf.fit_transform(tfidf)

# Extract topics
def extract_topic_keywords(model, feature_names, n_top_words=10):
    topics = []
    for topic_idx, topic in enumerate(model.components_):
        top_words = [feature_names[i] for i in topic.argsort()[:-n_top_words-1:-1]]
        topics.append(top_words)
    return topics

topics = extract_topic_keywords(nmf, feature_names)
topic_names = [generate_topic_name(topic_keywords) for topic_keywords in topics]

# Assign categories and urgency levels
df['Category'] = [topic_names[vec.argmax()] for vec in nmf_output]
df[['Urgency_Level', 'Urgency_Reason']] = pd.DataFrame([
    determine_urgency(comment) for comment in df['comments']
])

# Add confidence scores and key terms
topic_probabilities = nmf_output / nmf_output.sum(axis=1, keepdims=True)
df['Confidence_Score'] = [max(probs) for probs in topic_probabilities]
df['Alternative_Category'] = [topic_names[probs.argsort()[-2]] for probs in topic_probabilities]

def extract_key_terms(comment):
    """Extract urgent and topical key terms"""
    comment_lower = comment.lower()
    urgent_terms = []
    
    # Check for urgency keywords
    for level, rules in URGENCY_RULES.items():
        for keyword in rules['keywords']:
            if keyword in comment_lower:
                urgent_terms.append(keyword)
    
    return ', '.join(set(urgent_terms))

df['Key_Terms'] = df['comments'].apply(extract_key_terms)

# Calculate response timeframe
df['Response_Required_Within'] = df['Urgency_Level'].map({
    'RED_ALERT': '24 hours',
    'HIGH': '24 hours',
    'MEDIUM': '48 hours',
    'NORMAL': '72 hours'
})

# Save detailed results
output_columns = [
    'comments', 'Category', 'Urgency_Level', 'Urgency_Reason',
    'Response_Required_Within', 'Confidence_Score', 'Alternative_Category',
    'Key_Terms'
]
df[output_columns].to_excel("urgent_categorized_comments.xlsx", index=False, engine='openpyxl')

# Generate summary statistics
urgency_summary = df.groupby(['Urgency_Level', 'Category']).size().unstack(fill_value=0)
urgency_summary.to_excel("urgency_summary.xlsx")

# Print summary of urgent cases
print("\nUrgent Cases Summary:")
print(f"Red Alert Cases: {sum(df['Urgency_Level'] == 'RED_ALERT')}")
print(f"High Priority Cases: {sum(df['Urgency_Level'] == 'HIGH')}")
print(f"Medium Priority Cases: {sum(df['Urgency_Level'] == 'MEDIUM')}")