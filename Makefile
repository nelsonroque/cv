QUARTO ?= quarto

.PHONY: render
render:
	python3 scripts/generate_career_highlights.py
	$(QUARTO) render nelson-roque-tenure-cv.qmd --to pdf --output-dir rendered
	cp rendered/nelson-roque-tenure-cv.pdf /tmp/nelson-roque-tenure-cv.pdf
	$(QUARTO) render nelson-roque-tenure-cv.qmd --to docx --output-dir rendered
	cp /tmp/nelson-roque-tenure-cv.pdf rendered/nelson-roque-tenure-cv.pdf

p: render

ch:
	python3 scripts/generate_career_highlights.py