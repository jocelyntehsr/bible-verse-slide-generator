import pandas as pd
from constants import bible_books
import re

# ====== Load the Cleaned File ======
df = pd.read_csv("bible-verse-slide-generator\cuv_clean.csv")

# ====== Helper: Parse Reference ======
def parse_reference(ref: str):
    """
    Example input: "创世记 1:1-3"
    Returns: (book_name, chapter, [verses])
    """
    book, rest = ref.split(" ", 1)
    chapter, verses = rest.split(":")
    chapter = int(chapter)

    if "-" in verses:
        start, end = map(int, verses.split("-"))
        verse_list = list(range(start, end + 1))
    else:
        verse_list = [int(verses)]

    return book, chapter, verse_list


# ====== Query Function ======
# def query_bible(reference: str):

#     book, chapter, verses = parse_reference(reference)

#     subset = df[
#         (df["Book Name"] == book) &
#         (df["Chapter"] == chapter) &
#         (df["Verse"].isin(verses))
#     ][["Book Name", "Chapter", "Verse", "CleanText"]]

#     return subset

def query_sermon_bible(reference: str):

    book, chapter, verses = parse_reference(reference)

    subset = df[
        (df["Book Name"] == book) &
        (df["Chapter"] == chapter) &
        (df["Verse"].isin(verses))
    ][["Book Name", "Chapter", "Verse", "CleanText"]]

    subset["Title (CN)"] = reference

    pattern = r"^(\S+)\s+(.+)$" 
    match = re.match(pattern, reference)
    if match:
        book_cn, chapter_verse = match.groups()
        book_en = bible_books.get(book_cn, book_cn)  # fallback to original if not found
        result = f"{book_en} {chapter_verse}"

        subset["Title (EN)"] = result
        subset["Book Name (EN)"] = book_en

    return subset

# ====== Example Usage ======
if __name__ == "__main__":
    ref = ["约翰福音 10:11-15",
    "彼得前书 5:2-11"]
    verses_list = []

    for i in ref:
        verses = query_sermon_bible(i)
        verses_list.append(verses)

    combined_df = pd.concat(verses_list, ignore_index=True)
    combined_df.to_csv("bible_verses.csv", index=False, encoding="utf-8-sig")
    print(f"✅ Saved in csv")