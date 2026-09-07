"""Frozen OfficeBench evaluator excerpt used by task 2-20/0 at commit b978b808.

This audit copy preserves the upstream semantics. It is not imported by the
CoworkEvoBench Judge because keyword containment is insufficient for v0.3.
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
        if keyword not in content:
            return False
    return True


def evaluate_contain(testbed_dir, args):
    """Upstream dispatch reads args['file'] by doc_type, then applies the helper.

    The original function depends on OfficeBench application readers; this
    frozen excerpt records the exact containment decision used for this seed.
    """
    raise RuntimeError("Audit-only excerpt: use the frozen OfficeBench checkout for dispatch")
