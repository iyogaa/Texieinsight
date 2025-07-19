import streamlit as st
import pandas as pd
import re
import os
import numpy as np
from fuzzywuzzy import process, fuzz
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import Pipeline
import joblib
from sklearn.preprocessing import LabelEncoder

def mvr_gpt_app():
    st.markdown('<div class="custom-heading">MVR GPT Tool</div>', unsafe_allow_html=True)
    st.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTIGEqSiDgNs2c5VkcZ9eUba_LVjvy74f7w-w&s",
             width=100, caption="")

    @st.cache_data
    def load_data():
        try:
            df = pd.read_excel("Violation GPT MODEL.xlsx", engine="openpyxl")
            df.columns = df.columns.str.strip()
            df = df.dropna(subset=["Violation Description", "Category"])
            df["Violation Description"] = df["Violation Description"].str.strip()
            return df
        except Exception as e:
            st.error(f"❌ Failed to load data: {e}")
            return pd.DataFrame()

    df = load_data()
    if df.empty:
        st.stop()

    X_text = df["Violation Description"].str.lower()
    y = df["Category"]

    label_encoder = LabelEncoder()
    y_encoded = label_encoder.fit_transform(y)

    tfidf = TfidfVectorizer(ngram_range=(1, 2), stop_words="english")
    X_vect = tfidf.fit_transform(X_text)

    def train_and_save_model(X_vect, y_encoded, tfidf, label_encoder, model_path):
        model = Pipeline([("clf", LogisticRegression(max_iter=500))])
        model.fit(X_vect, y_encoded)
        joblib.dump((model, tfidf, label_encoder), model_path)
        return model, tfidf, label_encoder

    model_path = "violation_model.pkl"
    if not os.path.exists(model_path):
        st.info("Training model for the first time...")
        model, tfidf, label_encoder = train_and_save_model(X_vect, y_encoded, tfidf, label_encoder, model_path)
    else:
        loaded_obj = joblib.load(model_path)
        if isinstance(loaded_obj, tuple) and len(loaded_obj) == 3:
            model, tfidf, label_encoder = loaded_obj
        else:
            st.warning("⚠️ Existing pickle file is invalid. Retraining model...")
            model, tfidf, label_encoder = train_and_save_model(X_vect, y_encoded, tfidf, label_encoder, model_path)

    non_moving_keywords = [
        "improper equipment", "defective equipment", "traffic fines", "penalties",
        "lic", "fine", "court", "suspension", "misc", "sticker", "tags", "miscellaneous",
        "background check", "notice", "seat belt", "insurance", "certificate",
        "weighing", "loading", "length", "carrying", "loads", "susp", "seatbelt",
        "failure to signal", "illegal stop", "obstructing traffic","law"
    ]
    non_moving_keywords = [kw.lower() for kw in non_moving_keywords]

    rules = {
        "Accident Violation": ["collision", "crash", "hit and run"],
        "Major Violation": ["reckless", "dui", "excessive speeding", "dangerous"],
        "Prohibited Violation": ["prohibited", "unauthorized", "restricted"],
        "Minor Violation": ["speeding", "late payment", "parking violation"]
    }

    def detect_priority(desc):
        desc = desc.lower()
        if any(kw in desc for kw in non_moving_keywords):
            return "🚨 **Non-Moving Violation**"
        match = re.search(r"(\d{2,})/(\d{2,})", desc)
        if match:
            num, denom = map(int, match.groups())
            if num < denom:
                return "Minor Violation"
            elif num - denom >= 20:
                return "🚨 Major Violation"
            else:
                return "⚠️ Minor Violation"
        for lbl, kw_list in rules.items():
            if any(k in desc for k in kw_list):
                return f"🚨 Rule-Based: **{lbl}**"
        return "Unknown Violation"

    def classify_violation(description):
        desc = description.strip().lower()
        exact = df[df["Violation Description"].str.lower() == desc]
        if not exact.empty:
            return f"✅ Exact: **{exact['Category'].values[0]}**"
        rule = detect_priority(desc)
        if rule != "Unknown Violation,Better ask QA Team":
            return rule
        vec = tfidf.transform([desc])
        proba = model.predict_proba(vec)
        idx = np.argmax(proba)
        predicted_label = label_encoder.inverse_transform([idx])[0]
        confidence = proba[0][idx] * 100
        return f"🤖 Partial Prediction: **{predicted_label}** (Confidence: {confidence:.2f}%)"

    user_input = st.text_input("🔍 Enter Violation Description:")
    if user_input:
        if user_input.strip().lower() in ["yogaraj", "yoga"]:
            st.success("🐉 **Dragon Warrior** 🐼")
        else:
            result = classify_violation(user_input)
            st.info(result)