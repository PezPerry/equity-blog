#!/usr/bin/env python3
"""Company data quizzes: check every quiz file and rebuild quizzes/index.json.

One quiz per company, in quizzes/<slug>.json. The app's /Quizzes pages and the
front page read index.json, which this script writes - never edit it by hand.

    python quizzes/build-quizzes.py             check all quizzes, rebuild index.json
    python quizzes/build-quizzes.py --check     check only, write nothing
    python quizzes/build-quizzes.py --pending   list assessments with no current quiz

A quiz that fails a check stops the build (exit 1) and index.json is left alone,
so a bad file can never reach the site. Rules are in quizzes/README.md.
"""
import argparse
import datetime
import html
import importlib.util
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
INDEX = os.path.join(HERE, "index.json")

QUESTIONS = 8
TOPICS = {"profit", "margins", "cash", "returns", "sales", "debt", "dividends", "business", "guidance", "valuation"}
MONEY_TOPICS = {"profit", "margins", "cash", "returns"}   # "does it make money, and how much"
MIN_MONEY = 4
VERDICT_WORDS = re.compile(
    r"\b(strong buy|buy|hold|sell|short|reduce|avoid|overvalued|undervalued|fair value|verdict|price target|"
    r"target price|recommendation)\b", re.I)
SLUG_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")


def front_page():
    spec = importlib.util.spec_from_file_location("bfp", os.path.join(REPO, "build-front-page.py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def squash(text):
    """Lower-case, typographic marks flattened, whitespace collapsed - for quote matching."""
    text = html.unescape(text)
    for a, b in (("‘", "'"), ("’", "'"), ("“", '"'), ("”", '"'), ("–", "-"),
                 ("—", "-"), ("−", "-"), (" ", " ")):
        text = text.replace(a, b)
    return re.sub(r"\s+", " ", text).strip().lower()


def article_text(stem):
    path = os.path.join(REPO, stem + ".html")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as fh:
        page = fh.read()
    page = re.sub(r"<(script|style)\b.*?</\1>", " ", page, flags=re.S | re.I)
    return squash(re.sub(r"<[^>]+>", " ", page))


def assessments():
    """Latest published assessment per company: {TICKER: {...}}."""
    bfp = front_page()
    with open(os.path.join(REPO, "front-page-tiles.json"), encoding="utf-8") as fh:
        data = json.load(fh)
    tiles = data["tiles"] if isinstance(data, dict) else data
    latest = {}
    for tile in sorted(tiles, key=lambda t: t["date"]):
        info = bfp.assessment(tile)
        if not info or not info.get("ticker"):
            continue
        ticker = info["ticker"].rstrip(".").upper()
        latest[ticker] = {"ticker": ticker, "assessment": tile["url"][:-5], "date": tile["date"],
                          "headline": tile["headline"], "file": tile["url"]}
    return latest


def check_quiz(path, latest):
    errors = []
    name = os.path.basename(path)
    try:
        with open(path, encoding="utf-8") as fh:
            quiz = json.load(fh)
    except (OSError, ValueError) as exc:
        return None, ["%s: not valid JSON (%s)" % (name, exc)]

    def need(key, kind=str):
        value = quiz.get(key)
        if not isinstance(value, kind) or (kind is str and not value.strip()):
            errors.append("%s: '%s' is missing or empty" % (name, key))
        return value

    slug = need("slug")
    for key in ("company", "ticker", "title", "blurb", "assessment", "assessment_title",
                "assessment_date", "published"):
        need(key)
    if not isinstance(quiz.get("sector", ""), str):
        errors.append("%s: 'sector' must be text (it may be empty)" % name)
    if errors:
        return None, errors
    if not SLUG_RE.match(slug) or slug + ".json" != name:
        errors.append("%s: slug must be lower-case-with-hyphens and match the file name" % name)
    ticker = quiz["ticker"].rstrip(".").upper()
    current = latest.get(ticker)
    if not current:
        errors.append("%s: ticker %s has no published assessment" % (name, ticker))
    elif current["assessment"] != quiz["assessment"]:
        errors.append("%s: built from '%s' but the latest %s assessment is '%s' - replace the quiz"
                      % (name, quiz["assessment"], ticker, current["assessment"]))
    text = article_text(quiz["assessment"])
    if text is None:
        errors.append("%s: assessment file %s.html is not in the repo" % (name, quiz["assessment"]))
    cover = quiz.get("cover", "")
    if cover and not os.path.exists(os.path.join(REPO, cover)):
        errors.append("%s: cover %s does not exist" % (name, cover))

    questions = quiz.get("questions")
    if not isinstance(questions, list) or len(questions) != QUESTIONS:
        errors.append("%s: needs exactly %d questions" % (name, QUESTIONS))
        return None, errors
    answers, kinds, money = [], [], 0
    for number, q in enumerate(questions, start=1):
        where = "%s q%d" % (name, number)
        if q.get("id") != number:
            errors.append("%s: id must be %d" % (where, number))
        kind = q.get("type")
        options = q.get("options")
        if kind == "choice":
            if not (isinstance(options, list) and len(options) == 4 and len(set(options)) == 4):
                errors.append("%s: a choice question needs four different options" % where)
        elif kind == "higher_lower":
            if options != ["Higher", "Lower"]:
                errors.append('%s: a higher_lower question needs options ["Higher", "Lower"]' % where)
        else:
            errors.append("%s: type must be 'choice' or 'higher_lower'" % where)
        kinds.append(kind)
        answer = q.get("answer")
        if not isinstance(answer, int) or isinstance(answer, bool) or not isinstance(options, list) \
                or not 0 <= answer < len(options):
            errors.append("%s: answer must be the position of the right option, counting from 0" % where)
        else:
            answers.append((kind, answer))
        if q.get("topic") not in TOPICS:
            errors.append("%s: topic must be one of %s" % (where, ", ".join(sorted(TOPICS))))
        elif q["topic"] in MONEY_TOPICS:
            money += 1
        for key, low, high in (("prompt", 20, 240), ("explanation", 40, 420)):
            value = q.get(key, "")
            if not isinstance(value, str) or not low <= len(value.strip()) <= high:
                errors.append("%s: %s must be %d to %d characters" % (where, key, low, high))
        quote = q.get("source_quote", "")
        if not isinstance(quote, str) or len(quote.strip()) < 25:
            errors.append("%s: source_quote must be a sentence copied from the assessment" % where)
        elif text is not None and squash(quote) not in text:
            errors.append("%s: source_quote is not in the assessment word for word" % where)
        shown = " ".join([str(q.get("prompt", "")), str(q.get("explanation", ""))] + [str(o) for o in options or []])
        if VERDICT_WORDS.search(shown):
            errors.append("%s: mentions a verdict word (%s)" % (where, VERDICT_WORDS.search(shown).group(0)))
        if "~" in shown:
            errors.append("%s: contains a tilde" % where)
    if not 1 <= kinds.count("higher_lower") <= 2:
        errors.append("%s: needs one or two higher_lower questions" % name)
    if money < MIN_MONEY:
        errors.append("%s: at least %d questions must be about profit, margins, cash or returns" % (name, MIN_MONEY))
    spots = [a for kind, a in answers if kind == "choice"]
    if spots and max(spots.count(n) for n in set(spots)) > 3:
        errors.append("%s: the right answer sits in the same position more than three times" % name)
    for key in ("title", "blurb"):
        if VERDICT_WORDS.search(quiz[key]) or "~" in quiz[key]:
            errors.append("%s: %s mentions a verdict word or has a tilde" % (name, key))
    return quiz, errors


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--pending", action="store_true")
    args = parser.parse_args()

    latest = assessments()
    files = sorted(f for f in os.listdir(HERE) if f.endswith(".json") and f != "index.json")

    if args.pending:
        have = {}
        for name in files:
            try:
                with open(os.path.join(HERE, name), encoding="utf-8") as fh:
                    quiz = json.load(fh)
                have[quiz["ticker"].rstrip(".").upper()] = quiz.get("assessment")
            except (OSError, ValueError, KeyError):
                continue
        pending = [a for t, a in sorted(latest.items(), key=lambda kv: kv[1]["date"]) if have.get(t) != a["assessment"]]
        print(json.dumps(pending, indent=2, ensure_ascii=False))
        return 0

    quizzes, problems, tickers = [], [], {}
    for name in files:
        quiz, errors = check_quiz(os.path.join(HERE, name), latest)
        problems += errors
        if quiz and not errors:
            ticker = quiz["ticker"].rstrip(".").upper()
            if ticker in tickers:
                problems.append("%s and %s are both for %s - one quiz per company" % (tickers[ticker], name, ticker))
            tickers[ticker] = name
            quizzes.append(quiz)
    if problems:
        print("QUIZ CHECK FAILED - index.json not changed:")
        for problem in problems:
            print("  - " + problem)
        return 1

    quizzes.sort(key=lambda q: (q["published"], q["slug"]), reverse=True)
    index = {"quizzes": [
        {key: quiz.get(key, "") for key in ("slug", "company", "ticker", "sector", "title", "blurb", "cover",
                                             "assessment", "assessment_title", "assessment_date", "published")}
        | {"questions": len(quiz["questions"])} for quiz in quizzes]}
    body = json.dumps(index["quizzes"], ensure_ascii=False, sort_keys=True)
    old = {}
    if os.path.exists(INDEX):
        with open(INDEX, encoding="utf-8") as fh:
            old = json.load(fh)
    if json.dumps(old.get("quizzes"), ensure_ascii=False, sort_keys=True) == body:
        print("%d quizzes pass; index.json already up to date" % len(quizzes))
        return 0
    if args.check:
        print("%d quizzes pass; index.json would change (check only)" % len(quizzes))
        return 0
    index = {"updated": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"), **index}
    with open(INDEX, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(index, fh, ensure_ascii=False, indent=1)
        fh.write("\n")
    print("%d quizzes pass; index.json written" % len(quizzes))
    return 0


if __name__ == "__main__":
    sys.exit(main())
