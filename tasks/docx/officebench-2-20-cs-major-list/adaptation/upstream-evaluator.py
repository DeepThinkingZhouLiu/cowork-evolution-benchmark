"""Frozen OfficeBench evaluator excerpt used by task 2-20/1 at commit b978b808.

This audit copy preserves the upstream containment semantics. It is not
imported by the v0.3 Judge because keyword presence is not semantic evidence.
"""


def _is_number(string):
    try:
        float(string)
        return True
    except ValueError:
        return False


def _evaluate_contain_text(content, args):
    content = content.lower()
    for keyword in args["keywords"]:
        keyword = keyword.lower()
        if _is_number(keyword):
            content = content.replace(",", "")
        if keyword in content:
            continue
        return False
    return True


def evaluate_contain(testbed_dir, args):
    """Upstream dispatch reads args['file'] by doc_type, then applies the helper.

    The original function depends on OfficeBench application readers; use the
    frozen OfficeBench checkout to execute dispatch. This excerpt is audit-only.
    """
    raise RuntimeError("Audit-only excerpt: use the frozen OfficeBench checkout for dispatch")
