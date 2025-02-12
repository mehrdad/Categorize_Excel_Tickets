import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
import re
from collections import defaultdict

# Load comments from Excel with proper column name
df = pd.read_excel("comments.xlsx", engine='openpyxl')
comments = df['comments'].tolist()

# Enhanced preprocessing with key phrase preservation
def preprocess(text):
    # Key phrases to preserve
    key_phrases = [
        "not working", "inverter", "app issue", "portal", "battery", "roof leak",
        "installer", "payment deferment", "refund", "buyout", "warranty",
        "no ticket", "with ticket", "cancellation", "roof damage"
    ]
    
    text = str(text).lower()
    
    # Preserve key phrases
    for phrase in key_phrases:
        if phrase in text:
            text = text.replace(phrase, phrase.replace(" ", "_"))
    
    # Remove special characters but keep underscores
    text = re.sub(r'[^a-zA-Z0-9\s_]', '', text)
    return text.strip()

# Category-specific keyword mappings
category_keywords = {
    "System Malfunction -  Inverter": ["inverter", "not_working", "inverter_not", "inverter_issue"],
    "System Portal/ App Issue": ["app", "portal", "not_displaying", "not_communicating"],
    "System (Panel) Not Working - with ticket": ["ticket", "case", "not_working", "service_request"],
    "System (Panel) Not Working - without ticket": ["not_working", "no_power", "not_producing"],
    "System (Panel) Not Working - with request for payment deferment": ["payment", "bill", "not_working", "deferment"],
    "Installation Concerns - Installer Issue": ["installer", "installation", "damage", "incorrect"],
    "Installation Concerns - Roof Leaks": ["leak", "roof", "water", "damage"],
    "Rebate/ Refund/ Reimbursement Request": ["refund", "reimburse", "credit", "money_back"],
    "System Buy-out Procedure": ["buyout", "purchase", "own", "take_over"],
    "Customer Experience Issues - Request for Call/ Follow up": ["call_back", "contact", "follow_up", "reach"],
}

# Enhanced categorization function
def categorize_comment(comment):
    processed = preprocess(comment)
    scores = defaultdict(int)
    
    # Score each category based on keyword matches
    for category, keywords in category_keywords.items():
        for keyword in keywords:
            if keyword in processed:
                scores[category] += 1
                
                # Additional scoring rules
                if category.startswith("System (Panel) Not Working"):
                    if "ticket" in processed or "case" in processed:
                        scores["System (Panel) Not Working - with ticket"] += 1
                    elif "payment" in processed or "bill" in processed:
                        scores["System (Panel) Not Working - with request for payment deferment"] += 1
                    else:
                        scores["System (Panel) Not Working - without ticket"] += 1
                
    # Return category with highest score, or uncategorized if no matches
    if scores:
        return max(scores.items(), key=lambda x: x[1])[0]
    return "Uncategorized"

# Process the comments
processed_comments = [preprocess(comment) for comment in comments]

# Generate TF-IDF matrix with improved parameters
vectorizer = TfidfVectorizer(
    max_features=1000,
    stop_words='english',
    ngram_range=(1, 2),  # Include bigrams
    max_df=0.9,
    min_df=2
)
tfidf_matrix = vectorizer.fit_transform(processed_comments)

# Cluster with improved parameters
n_clusters = len(category_keywords)  # Match clusters to number of main categories
kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
clusters = kmeans.fit_predict(tfidf_matrix)

# Extract and map keywords
feature_names = vectorizer.get_feature_names_out()
cluster_keywords = defaultdict(list)
for i in range(n_clusters):
    cluster_indices = clusters == i
    cluster_tfidf = tfidf_matrix[cluster_indices].mean(axis=0).A1
    top_keyword_indices = cluster_tfidf.argsort()[-10:][::-1]
    keywords = [feature_names[idx] for idx in top_keyword_indices]
    cluster_keywords[f"Cluster_{i}"] = keywords

# Categorize comments and save results
df['Category'] = df['comments'].apply(categorize_comment)
df.to_excel("auto1_categorized_comments.xlsx", index=False, engine='openpyxl')