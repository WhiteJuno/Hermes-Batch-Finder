The search engine follows a strict "priority-based" selection process to find the single match.

1.- Client ID Filter
    Before any matching happens, the tool ignores everything in the database that doesn't belong to the Client ID you selected (e.g., EMEA,           Canada). Th  is ensures a batch with the same name from a different region is never shown.

2.- Search "Hierarchy"
    The tool then tries two ways to find a match, in this specific order:
    Priority #1: The Exact Match
      It first checks if the input matches any batch name in the database character-for-character (ignoring capitalization).
      If it finds an exact match, it immediately stops and picks that row.
    Priority #2: The Weighted Similarity Score (Fuzzy Match)
      If there is no exact match, the tool calculates a Similarity Score for every batch in the filtered list. It then picks the row with the           highest score. This score is a blend of three different mathematical tests:

Test A: Longest Shared Fragment (40% of the score): 
      Hermes looks for the longest "unbroken block" of text that exists in both your input and the database entry. This is very good at                 finding a batch name even if it's buried in a long string of pasted text.
      
      This measures the longest contiguous piece of text that appears in both strings.
      $$ \text{Score A} = \frac{\text{Length of Longest Common Substring}}{\text{maxLen}} $$
      Example: If your input is BE_BIP (6 chars) and the database is BE_BIP_BUS (10 chars), the longest shared part is BE_BIP (6 chars), the            score is 6 / 10 = 0.60.
      
Test B: Word/Term Overlap (Jaccard Index, 30% of the score):
      Hermes breaks both your input and the database batch name into individual "words" (splitting at _, -, /, and spaces). It counts how many of       these unique words match. This helps it ignore extra dates or random numbers at the end of a pasted batch.

      The engine splits both strings into "words" at characters like spaces, underscores (_), dashes (-), and slashes (/).
      $$ \text{Score B} = \frac{\text{Number of Unique Common Words}}{\text{Total Number of Unique Words across Both Strings}} $$
      Example:
      Input: BE, BIP, 2025 (3 unique words)
      Database: BE, BIP, BUS (3 unique words)
      The common words are BE and BIP (2 words).
      The total set of unique words is BE, BIP, BUS, 2025 (4 words).
      The score is 2 / 4 = 0.50.
      
Test C: Overall Resemblance / Edit Distance (Normalized Levenshtein, 30% of the score):
      It calculates how many "edits" (inserting, deleting, or changing a letter) would be needed to turn your input into the database version.

      The Levenshtein $(\textit{lev})$ measures the minimum number of single-character edits (insertions, deletions, or substitutions) required         to change one string into the other.
      $$ \text{Score C} = 1 - \left( \frac{\text{lev(Input, Database)}}{\text{maxLen}} \right) $$
      Example: If it takes 2 edits to turn your input into the database version and the longest string is 10 characters long, the score is              1 - (2 / 10) = 0.80.

Why this works:
By combining these three tests, the "winner" is usually the row that shares the most meaningful parts of the batch name, even if the analyst who pasted it included extra words, dates or numbers that aren't in the database.

      The three scores are then weighted and added together to find the Grand Total. The database entry with the highest Grand Total wins:
      $$ \text{Grand Total Score} = ( \text{Score A} \times 0.40 ) + ( \text{Score B} \times 0.30 ) + ( \text{Score C} \times 0.30 ) $$
      This means that having a long shared fragment (Test A) is the most powerful factor in deciding which batch is the "Winner."

-----------------------------------------------------------------------------

The search engine ONLY shows the results of the single batch that was selected as the closest or exact match. It does not combine, fuse, or concatenate strings from different rows.

1.- Isolation: First, it grabs all batches belonging to the selected Client ID.
2.- If there is an Exact Match, it picks that specific row and stops looking.
3.- If there is no exact match (like most of the times), it calculates a "similarity score" for every batch under that Client ID and identifies       the single row with the highest score.
4.- Display: Once it has that one "Selected" row (stored in the variable bestMatch), it looks only at that row's data.
    row[2] is mapped directly to the Strong Potential Match card.
    row[3] is mapped directly to the Pending / Insufficient card.

You will never see a mixed result where the SPM comes from Batch A and the Pending/Insuff comes from Batch B. Every piece of information on the screen belongs to the same unique entry as it was written in the Excel database.

-----------------------------------------------------------------------------

For the Excel formula, we use:
="[""" & B2 & """,""" & C2 & """,""" & D2 & """,""" & E2 & """],"
to format the database and paste it into the code.
