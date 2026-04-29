F1 Commentary Nationality & Sentiment Analysis
This repository contains the Python script developed for an AP Research study examining nationalistic bias in Formula 1 broadcast commentary. The script analyzes transcripts from F1 Live and Sky Sports commentary across two races from the 2025 Formula 1 season — the British Grand Prix and the Australian Grand Prix.
What the script does:

Identifies and counts mentions of each driver by name across all transcripts
Identifies and counts explicit nationality references attached to driver mentions
Applies AFINN lexicon-based sentiment scoring to language used when describing drivers
Excludes racing-specific terminology from sentiment scoring using a custom dictionary to reduce false positives

Data sources:

Commentary transcripts were obtained from F1 Live (F1 TV) and Sky Sports F1 broadcasts of the 2025 British Grand Prix and 2025 Australian Grand Prix
Transcripts were generated using Otter.AI and manually reviewed and corrected by the researcher

Tools and dependencies:

Python 3
AFINN lexicon (Nielsen, 2011)
Claude (Anthropic) was used to support code development and debugging; all coding logic and final implementation were determined by the researcher

Note:
This script was developed for academic research purposes. Transcripts are not included in this repository due to copyright considerations.
