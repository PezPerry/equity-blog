# Company data quizzes: instructions for the quiz writer

This folder holds one quiz per company for Equity & Markets Insight. The site reads
`index.json`; the quiz player reads `<slug>.json`. You (Codex) write the quiz files and
covers. Two scripts check your work and build everything else. Decisions below were made by
Mickey on 12 November 2026.

## What a quiz is

- Eight questions on one company, written **only** from that company's latest published
  assessment on the site. Nothing from memory, the web, the company's accounts or any other
  article. If a fact is not in the assessment, it cannot be in the quiz.
- Pitched at a **beginner**. Someone who has never read a set of accounts should be able to
  attempt every question, and learn something from every answer.
- A stand-alone piece. The reader may never have seen the assessment, so each question must
  make sense on its own: say what a term means inside the question the first time it appears
  ("Revenue is the money taken in from customers before any costs. How much did ...").
- **Weighted towards whether the company makes money and how much.** At least four of the
  eight questions must be about profit, margins, cash or returns. The rest can cover anything
  else the assessment says: sales, debt, dividends, what the business does, guidance.
- One quiz per company. When a company gets a new assessment, its quiz is **replaced**, not
  added to: same file name, new contents.

## What a quiz must never do

- Never state, ask about or hint at the Buy, Hold or Sell call, the valuation, a target price
  or the "value opportunity". Do not use the words buy, hold, sell, short, reduce, avoid,
  overvalued, undervalued, verdict or recommendation anywhere in a quiz, even innocently
  ("sells watches" is fine; "sell" as a word on its own is rejected by the checker - reword).
- No share price predictions and nothing about what the shares will do.
- No trick questions: no double negatives, no "all of the above", no two answers that are both
  defensible, no questions that turn on a rounding difference.
- No trivia about people (names of directors, founders' birthdays).
- No opinion presented as fact. Ask about what the assessment reports, not whether it is good.
- No emoji. British English. Never put a tilde before a number; write "around 28%".
- Do not copy the assessment's first-person opinions into explanations ("I think ..."). The
  explanation states the fact and why it matters in plain words.

## Format

Each quiz has exactly eight questions:

- Six or seven `choice` questions: four different options, one right.
- One or two `higher_lower` questions: options exactly `["Higher", "Lower"]`. Give the
  comparison figure in the question ("A year earlier net debt was £112m. Was it higher or
  lower at ...?").
- Wrong options must be plausible and of the same kind as the right one (all percentages, all
  pound figures). Where the assessment gives other real figures, use those as wrong options.
- Spread the right answers around: the checker rejects a quiz where the right answer sits in
  the same position more than three times.
- `prompt`: 20 to 240 characters. `explanation`: 40 to 420 characters, two or three short
  sentences: the right answer, then why it matters to someone judging whether the business
  makes money. The explanation is shown straight after the reader answers.
- `source_quote`: one sentence or clause copied **word for word** from the assessment that
  proves the right answer. The checker searches the published article for it and rejects the
  quiz if it is not there. It is never shown to the reader. Quote running prose, not table
  cells.
- `topic`: one of `profit`, `margins`, `cash`, `returns`, `sales`, `debt`, `dividends`,
  `business`, `guidance`, `valuation`. (Use `valuation` only for neutral facts such as the
  number of shares in issue, never for the call.)

Order the questions so the quiz tells a small story: what the business does, what it takes in,
what it keeps, what cash it produces, what it owes.

`watches-of-switzerland-wosg.json` in this folder is the worked example. Match its tone.

## The file

`quizzes/<slug>.json`, where `<slug>` is the company's short name and ticker in lower case
with hyphens, for example `greggs-grg`, `jd-sports-jd`, `bp-bp`. Use the same slug for that
company for ever.

```json
{
  "slug": "watches-of-switzerland-wosg",
  "company": "Watches of Switzerland",
  "ticker": "WOSG",
  "sector": "Consumer Discretionary",
  "title": "Watches of Switzerland: how well do you know the numbers?",
  "blurb": "One sentence on what the quiz covers.",
  "cover": "quizzes/covers/watches-of-switzerland-wosg.webp",
  "assessment": "watches-of-switzerland-group-plc-wosg-buy-running-well-priced-for-it",
  "assessment_title": "Running Well, Priced For It",
  "assessment_date": "2026-11-10",
  "published": "2026-11-12",
  "questions": [
    { "id": 1, "type": "choice", "topic": "business", "prompt": "...",
      "options": ["...", "...", "...", "..."], "answer": 1,
      "explanation": "...", "source_quote": "..." }
  ]
}
```

- `company`: the name a reader would use ("Greggs", not "Greggs plc").
- `title`: always "<Company>: how well do you know the numbers?".
- `sector`: the company's `industry` value in `company-data/data.js` (for example
  "Consumer Discretionary", "Financials"). Leave it as `""` if the company is not in that file.
- `assessment`: the article's file name without `.html`. `assessment_title`: the part of the
  headline after the colon, without the company, ticker or the call.
- `answer`: the position of the right option, counting from 0.
- `published`: today's date.

## The cover

Do not generate an image. Run:

```
python quizzes/make-cover.py --company "<logo library name>" --ticker <TICKER> --slug <slug>
```

`--company` is the file name, without `.json`, of the company's entry in
`C:\Master Control\System\Assets\Logos`. The script puts the approved logo on its approved
background above a navy "The numbers quiz" band. If the company has no logo entry the cover
shows the ticker instead and the script says so; that is acceptable, and the list of companies
without a logo goes in your run note. Never fetch or draw a logo yourself.

## Each run

Work in the publishing clone of `PezPerry/equity-blog` that the other pipelines use, after
`git pull --rebase`. Files in this repo use LF line endings.

1. `python quizzes/build-quizzes.py --pending` lists every assessment whose company has no
   quiz, or whose quiz was built from an older assessment. An empty list means nothing to do:
   stop.
2. For each one, read the article file it names, write the quiz (overwriting the company's
   old file if there is one), and make the cover.
3. `python quizzes/build-quizzes.py` checks every quiz and rewrites `index.json`. If it prints
   `QUIZ CHECK FAILED`, fix the quizzes it names and run it again. Never edit `index.json` by
   hand and never publish while the check fails.
4. `python build-front-page.py` refreshes the quizzes section of the front page.
5. Commit only `quizzes/` and `front-page.html`, then push. No approval is needed: Mickey has
   said the quiz writer publishes on its own.
6. Leave a short run note: which quizzes were written or replaced, which have ticker covers,
   and anything in an assessment that made a fair question hard to write.

Run once a day, early morning. A quiz for a new assessment should be live by the day after the
assessment. The first run will find every published assessment pending (27 on 12 November
2026): write them all in that run.

Company assessments only. Long reads, The Greats and Trading Log entries never get a quiz;
the pending list already leaves them out.

## How it is shown

One question per screen; the answer and explanation appear as soon as the reader chooses. At
the end the reader sees the score and a link to the assessment, nothing else. No timer, no
streaks, no leaderboard, no saved scores. Quizzes are free to everyone without signing in and
do not count towards the three free articles a month.
