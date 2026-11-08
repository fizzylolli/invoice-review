"""Quick checks that the flag rules do what the README says. Run: python test_rules.py"""
import unittest
from datetime import date
import invoice_pipeline as ip

D = date(2026, 11, 8)


def rec(file, vendor="Acme Co", num="A-1", total=100.0, inv_date=date(2026, 9, 1), **kw):
    r = dict(file=file, vendor=vendor, invoice_number=num, total=total, invoice_date=inv_date,
             due_final=date(2026, 10, 1), currency="USD", line_items=[total])
    r.update(kw)
    return r


def flags(*rows, prior=()):
    return ip.apply_flags(list(rows), D, prior)


class Rules(unittest.TestCase):
    def test_over_threshold(self):
        self.assertEqual(flags(rec("a", total=5000.01, line_items=[5000.01]))[0]["flag"], "Y")

    def test_exactly_at_threshold_is_not_flagged(self):
        self.assertEqual(flags(rec("a", total=5000.00, line_items=[5000.00]))[0]["flag"], "N")

    def test_clean_invoice(self):
        self.assertEqual(flags(rec("a"))[0]["flag"], "N")

    def test_duplicate_invoice_number_flags_both(self):
        a, b = flags(rec("a.pdf"), rec("b.pdf"))
        self.assertEqual((a["flag"], b["flag"]), ("Y", "Y"))
        self.assertIn("duplicate of a.pdf", b["reasons"][0])

    def test_duplicate_survives_ocr_zero_vs_letter_o(self):
        a, b = flags(rec("a", num="RW-1041"), rec("b", num="RW-1O41"))
        self.assertEqual(b["flag"], "Y")

    def test_duplicate_by_amount_and_date_when_number_differs(self):
        a, b = flags(rec("a", num="X-1"), rec("b", num="X-2"))
        self.assertEqual(b["flag"], "Y")

    def test_same_vendor_different_month_is_not_duplicate(self):
        a, b = flags(rec("a", num="N-1", total=3000, line_items=[3000]),
                     rec("b", num="N-2", total=3100, line_items=[3100], inv_date=date(2026, 10, 1)))
        self.assertEqual((a["flag"], b["flag"]), ("N", "N"))

    def test_line_items_do_not_add_up(self):
        r = flags(rec("a", total=1384.50, line_items=[485.00, 448.00, 351.50]))[0]
        self.assertEqual(r["flag"], "Y")
        self.assertIn("1,284.50", r["reasons"][0])

    def test_euro_invoice_flagged(self):
        r = flags(rec("a", total=4700, currency="EUR", line_items=[4700]))[0]
        self.assertEqual(r["flag"], "Y")
        self.assertIn("EUR", r["reasons"][0])

    def test_missing_field_flagged(self):
        r = flags(rec("a", num=None))[0]
        self.assertEqual(r["flag"], "Y")
        self.assertIn("could not read", r["reasons"][0])

    def test_portfolio_company_flagged(self):
        r = flags(rec("a", text="placement fee for grasp technologies"))[0]
        self.assertEqual(r["flag"], "Y")
        self.assertIn("Grasp", r["reasons"][0])

    def test_duplicate_of_a_previous_weeks_invoice(self):
        old = dict(file="old.pdf", vendor="Acme Co", invoice_number="A-1", total=100.0,
                   invoice_date=date(2026, 9, 1), prior=True)
        r = flags(rec("new.pdf"), prior=[old])[0]
        self.assertEqual(r["flag"], "Y")
        self.assertIn("earlier run", r["reasons"][0])


def with_rules(*rows, **changes):
    """Run flags with some rules changed, e.g. with_rules(r, threshold={"enabled": False})."""
    rules = ip.load_rules(None)
    for k, v in changes.items():
        if isinstance(v, dict):
            rules[k].update(v)
        else:
            rules[k] = v
    return ip.apply_flags(list(rows), D, (), rules)


class RuleSettings(unittest.TestCase):
    big = dict(total=9000.0, line_items=[9000.0])

    def test_turning_threshold_off(self):
        self.assertEqual(with_rules(rec("a", **self.big), threshold={"enabled": False})[0]["flag"], "N")

    def test_changing_threshold_amount(self):
        self.assertEqual(with_rules(rec("a", **self.big), threshold={"amount": 10000})[0]["flag"], "N")
        self.assertEqual(with_rules(rec("a", **self.big), threshold={"amount": 8000})[0]["flag"], "Y")

    def test_duplicate_rule_off(self):
        a, b = with_rules(rec("a"), rec("b"), duplicate={"enabled": False})
        self.assertEqual((a["flag"], b["flag"]), ("N", "N"))

    def test_overdue_is_off_by_default_and_works_when_on(self):
        self.assertEqual(flags(rec("a"))[0]["flag"], "N")            # due 2026-10-01, run 2026-11-08
        r = with_rules(rec("a"), overdue={"enabled": True, "days": 30})[0]
        self.assertEqual(r["flag"], "Y")
        self.assertIn("38 days past due", r["reasons"][0])

    def test_new_vendor_rule(self):
        r = with_rules(rec("a", vendor="Shady LLC"), new_vendor={"enabled": True, "known_vendors": ["Acme Co"]})[0]
        self.assertEqual(r["flag"], "Y")
        ok = with_rules(rec("a"), new_vendor={"enabled": True, "known_vendors": ["Acme Co"]})[0]
        self.assertEqual(ok["flag"], "N")

    def test_new_vendor_rule_does_nothing_with_empty_list(self):
        self.assertEqual(with_rules(rec("a"), new_vendor={"enabled": True, "known_vendors": []})[0]["flag"], "N")

    def test_custom_rule(self):
        rule = [{"name": "Legal", "enabled": True, "contains": ["legal services"],
                 "reason": "Legal invoice: confirm the matter is approved", "next_step": "Ask the COO.", "priority": 1}]
        r = with_rules(rec("a", text="legal services - august"), custom=rule)[0]
        self.assertEqual(r["flag"], "Y")
        self.assertEqual(ip.triage(r)[0], 1)
        self.assertEqual(with_rules(rec("a", text="paper and toner"), custom=rule)[0]["flag"], "N")

    def test_custom_rule_switched_off(self):
        rule = [{"name": "Legal", "enabled": False, "contains": ["legal"], "reason": "x"}]
        self.assertEqual(with_rules(rec("a", text="legal"), custom=rule)[0]["flag"], "N")

    def test_missing_settings_fall_back_to_defaults(self):
        import json, tempfile, os
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as f:
            json.dump({"threshold": {"amount": 123}}, f)
        rules = ip.load_rules(f.name); os.unlink(f.name)
        self.assertEqual(rules["threshold"], {"enabled": True, "amount": 123})
        self.assertTrue(rules["duplicate"]["enabled"])


class EditorInput(unittest.TestCase):
    def test_junk_from_the_browser_is_cleaned(self):
        import rules_editor
        r = rules_editor.clean_rules({"threshold": {"amount": "abc", "enabled": 1}, "custom": [5, {"name": "", "contains": "x"}],
                                      "overdue": {"days": "-3"}, "bogus": {"enabled": True}})
        self.assertEqual(r["threshold"]["amount"], 5000)           # bad number falls back
        self.assertEqual(r["overdue"]["days"], 0)                  # negative days clamped
        self.assertNotIn("bogus", r)
        self.assertEqual(len(r["custom"]), 1)                      # the non-dict entry is dropped
        self.assertEqual(r["custom"][0]["name"], "Untitled rule")

    def test_saved_rules_round_trip(self):
        import rules_editor
        r = rules_editor.clean_rules(ip.load_rules(None))
        self.assertEqual(r, ip.load_rules(None))


class AppWindow(unittest.TestCase):
    def test_folder_browser_lists_folders_and_counts_invoices(self):
        import rules_editor, tempfile, os
        with tempfile.TemporaryDirectory() as d:
            os.mkdir(os.path.join(d, "sub")); os.mkdir(os.path.join(d, ".hidden"))
            for n in ("a.pdf", "b.JPG", "notes.txt"):
                open(os.path.join(d, n), "w").close()
            r = rules_editor.list_dir(d)
            self.assertEqual([x["name"] for x in r["dirs"]], ["sub"])      # hidden folders skipped
            self.assertEqual(r["invoices"], 2)                              # txt is not an invoice

    def test_folder_browser_survives_a_bad_path(self):
        import rules_editor
        self.assertTrue(rules_editor.list_dir("/definitely/not/a/real/path")["dirs"] is not None)


class Parsing(unittest.TestCase):
    def test_dates(self):
        self.assertEqual(ip.parse_date("09/08/2026"), date(2026, 9, 8))        # US month/day
        self.assertEqual(ip.parse_date("19 September 2026"), date(2026, 9, 19))
        self.assertEqual(ip.parse_date("Sep 22, 2026"), date(2026, 9, 22))
        self.assertEqual(ip.parse_date("2026-10-24"), date(2026, 10, 24))

    def test_due_date_from_terms(self):
        r = ip.finish({"invoice_date": date(2026, 9, 4), "terms_days": 30})
        self.assertEqual(r["due_final"], date(2026, 10, 4))
        r = ip.finish({"invoice_date": date(2026, 9, 17), "due_on_receipt": True})
        self.assertEqual(r["due_final"], date(2026, 9, 17))

    def test_scanned_dollar_sign_misread_as_euro_stays_usd(self):
        text = "CLEARWATER COFFEE\nINVOICE NO. CW-4410\nDATE 09/17/2026\nCoffee 1x 186.00 186.00\nAMOUNT DUE €228.00\n"
        self.assertEqual(ip.local_extract([text], ocr=True)["currency"], "USD")

    def test_real_euro_invoice_in_text_layer_is_euro(self):
        text = "Fathom Research GmbH\nFR-2026-0412 · 09/29/2026\nTOTAL DUE €4,700.00\n"
        self.assertEqual(ip.local_extract([text], ocr=False)["currency"], "EUR")

    def test_scanned_euro_invoice_with_euro_words_is_euro(self):
        text = "Acme GmbH\nInvoice 77\nAll amounts in EUR\nTOTAL DUE €900.00\n"
        self.assertEqual(ip.local_extract([text], ocr=True)["currency"], "EUR")

    def test_extract_simple_invoice(self):
        text = ("SUMMIT OFFICE SUPPLY CO.\nINVOICE NO.   SOS-55910\nDATE          04 September 2026\n"
                "TERMS         Net 30\nCopy paper   10 x 48.50   485.00\nAMOUNT DUE   $485.00\n")
        r = ip.finish(ip.local_extract([text]))
        self.assertEqual((r["invoice_number"], r["total"], str(r["due_final"])), ("SOS-55910", 485.0, "2026-10-04"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
