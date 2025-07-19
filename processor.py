import os
import re
import time
from PyPDF2 import PdfReader
from pdf2image import convert_from_path
import pytesseract
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from collections import Counter
from PIL import Image

class PDFTextSearcher:
    def __init__(self):
        self.pages_text = []
        self.document_name = ""
        self.processed = False
        self.total_pages = 0
        self.page_images = []
        self.ocr_data = []

    def process_pdf(self, file_path):
        """Process PDF with optimized OCR and text extraction"""
        if self.processed:
            return

        self.document_name = os.path.basename(file_path)
        start_time = time.time()

        try:
            # First pass: try standard text extraction
            with open(file_path, 'rb') as f:
                reader = PdfReader(f)
                self.total_pages = len(reader.pages)
                self.pages_text = [''] * self.total_pages

                for i, page in enumerate(reader.pages):
                    text = page.extract_text()
                    if text and text.strip():
                        self.pages_text[i] = text.strip()

            # Always run OCR to get precise word positions for highlighting
            images = convert_from_path(file_path, dpi=250)
            self.page_images = images

            # Store OCR data for each page: (text, left, top, width, height)
            self.ocr_data = []
            for img in images:
                gray_img = img.convert('L')
                ocr_result = pytesseract.image_to_data(
                    gray_img,
                    output_type=pytesseract.Output.DICT
                )
                self.ocr_data.append(ocr_result)

                # Extract text from OCR
                text = " ".join([word for word in ocr_result['text'] if word.strip()])
                page_idx = images.index(img)
                if page_idx < len(self.pages_text) and len(text) > len(self.pages_text[page_idx]):
                    self.pages_text[page_idx] = text

            self.processed = True
            return f"✅ Processed {self.total_pages} pages in {time.time()-start_time:.2f}s"

        except Exception as e:
            return f"❌ Processing failed: {str(e)}"

    def find_phrase_positions(self, phrase, page_num):
        """Find precise positions of phrase using OCR data"""
        page_idx = page_num - 1
        if page_idx >= len(self.ocr_data):
            return []

        ocr_info = self.ocr_data[page_idx]
        words = ocr_info['text']

        # Normalize phrase and words
        norm_phrase = re.sub(r'\s+', ' ', phrase).strip().lower()
        norm_words = [word.lower().strip() for word in words]

        # Find all occurrences of the phrase
        matches = []
        for i in range(len(norm_words)):
            # Check for multi-word phrase match
            if norm_words[i] == norm_phrase.split()[0]:
                match = True
                for j in range(1, len(norm_phrase.split())):
                    if i+j >= len(norm_words) or norm_words[i+j] != norm_phrase.split()[j]:
                        match = False
                        break

                if match:
                    # Get combined bounding box for the phrase
                    x_vals = []
                    y_vals = []
                    widths = []
                    heights = []

                    for k in range(len(norm_phrase.split())):
                        idx = i + k
                        x_vals.append(ocr_info['left'][idx])
                        y_vals.append(ocr_info['top'][idx])
                        widths.append(ocr_info['width'][idx])
                        heights.append(ocr_info['height'][idx])

                    x_min = min(x_vals)
                    y_min = min(y_vals)
                    x_max = max(x + w for x, w in zip(x_vals, widths))
                    y_max = max(y + h for y, h in zip(y_vals, heights))

                    matches.append({
                        'x': x_min,
                        'y': y_min,
                        'w': x_max - x_min,
                        'h': y_max - y_min,
                        'text': " ".join(words[i:i+len(norm_phrase.split())])
                    })

        return matches

    def semantic_search(self, question):
        """Context-aware search with proximity scoring"""
        if not self.processed:
            return "Document not processed", [], 0.0, 0, None

        start_time = time.time()
        question_lower = question.lower()

        # Extract core question components
        question_keywords = self.extract_keywords(question_lower)
        if not question_keywords:
            return "Question too vague", [], 0.0, 0, None

        # Semantic scoring parameters
        best_score = 0
        best_match = None
        best_page = 0
        match_text = ""

        for page_num, text in enumerate(self.pages_text):
            text_lower = text.lower()

            # 1. Presence score - are all keywords present?
            presence_score = sum(1 for kw in question_keywords if kw in text_lower)
            if presence_score == 0:
                continue

            # 2. Density score - how close are keywords to each other?
            positions = {}
            for kw in question_keywords:
                start_pos = text_lower.find(kw)
                if start_pos != -1:
                    positions[kw] = start_pos

            if len(positions) < len(question_keywords):
                continue

            # Calculate proximity between keywords
            kw_positions = sorted(positions.values())
            distances = [kw_positions[i+1] - kw_positions[i] for i in range(len(kw_positions)-1)]
            avg_distance = sum(distances) / len(distances) if distances else 0

            # Density score is inverse to average distance
            density_score = 100 / (avg_distance + 1) if avg_distance > 0 else 100

            # 3. Context score - is there surrounding context that confirms?
            context_score = 0
            full_match = " ".join(question_keywords)
            if full_match in text_lower:
                context_score = 50
            elif any(qw in text_lower for qw in ["yes", "required", "performed"]):
                context_score = 30
            elif any(qw in text_lower for qw in ["no", "not required", "not performed"]):
                context_score = 30

            # 4. Positional boost - earlier pages get slight preference
            position_boost = (1 - (page_num / self.total_pages)) * 10

            # Total score
            total_score = presence_score * 20 + density_score + context_score + position_boost

            if total_score > best_score:
                best_score = total_score
                best_page = page_num

                # Find the most relevant sentence
                sentences = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', text)
                for sentence in sentences:
                    if all(kw in sentence.lower() for kw in question_keywords):
                        match_text = sentence.strip()
                        break

        if best_score > 0:
            # Normalize confidence (0-1 scale)
            confidence = min(1.0, best_score / 200)
            search_time = (time.time() - start_time) * 1000
            return match_text, [best_page + 1], confidence, search_time, question_keywords

        return "Information not found", [], 0.0, 0, None

    def extract_keywords(self, text):
        """Extract meaningful keywords from text"""
        stop_words = {"does", "do", "is", "are", "the", "a", "an", "what", "where", "how", "why", "when"}
        words = re.findall(r'\b\w{3,}\b', text.lower())
        return [word for word in words if word not in stop_words]

    def get_context(self, page_num, phrase, context_size=100):
        """Get text context around a phrase"""
        page_idx = page_num - 1
        if page_idx >= len(self.pages_text):
            return ""

        text = self.pages_text[page_idx]
        pos = text.lower().find(phrase.lower())

        if pos == -1:
            return f"Phrase not found on page {page_num}"

        # Extract surrounding context
        start_pos = max(0, pos - context_size)
        end_pos = min(len(text), pos + len(phrase) + context_size)

        context = text[start_pos:end_pos]
        hl_phrase = context[context.lower().find(phrase.lower()):context.lower().find(phrase.lower())+len(phrase)]
        context = context.replace(hl_phrase, f"**{hl_phrase}**")

        return context

    def visualize_page(self, page_num, highlight_phrases=None):
        """Display PDF page with precision phrase highlighting"""
        if not 1 <= page_num <= self.total_pages:
            return None, "Invalid page number"

        page_idx = page_num - 1
        if page_idx >= len(self.page_images) or page_idx >= len(self.ocr_data):
            return None, "Page data not available"

        try:
            img = self.page_images[page_idx]
            fig, ax = plt.subplots(figsize=(12, 16))
            ax.imshow(img)
            ax.axis('off')

            title = f"Page {page_num}"

            if highlight_phrases:
                colors = ['#FF4136', '#3D9970', '#FF851B', '#0074D9', '#B10DC9']
                color_idx = 0

                for phrase in highlight_phrases:
                    matches = self.find_phrase_positions(phrase, page_num)

                    for i, match in enumerate(matches):
                        rect = patches.Rectangle(
                            (match['x'], match['y']),
                            match['w'],
                            match['h'],
                            linewidth=2,
                            edgecolor=colors[color_idx % len(colors)],
                            facecolor=colors[color_idx % len(colors)] + '40',
                            alpha=0.7
                        )
                        ax.add_patch(rect)

                        plt.text(
                            match['x'],
                            match['y'] - 5,
                            f"Matches {i+1}",
                            color=colors[color_idx % len(colors)],
                            fontsize=12,
                            bbox=dict(facecolor='white', alpha=0.8, edgecolor='none', pad=1)
                        )

                    color_idx += 1

                title += f" | Highlighted: {len(highlight_phrases)} phrases"

            plt.title(title, fontsize=14)
            plt.tight_layout()
            return fig, ""
        except Exception as e:
            return None, f"Rendering error: {str(e)}"
