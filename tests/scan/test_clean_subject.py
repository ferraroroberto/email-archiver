"""Reply/forward prefix stripping must need the colon and keep real words whole."""
import pytest

from email_archiver.text import clean_subject


@pytest.mark.parametrize(
    "subject",
    [
        "Recibo de pago de Tesla",
        "Revisión del contrato",
        "Reunión de equipo",
        "Rvalue notes",
        "Fwdcheck results",
        "Fwy closure",
        "Re",
    ],
)
def test_word_starting_with_a_prefix_stays_whole(subject: str) -> None:
    assert clean_subject(subject) == subject


@pytest.mark.parametrize(
    ("subject", "expected"),
    [
        ("RE: Recibo de pago", "Recibo de pago"),
        ("re:Recibo de pago", "Recibo de pago"),
        ("Rv: Factura", "Factura"),
        ("Fw: Factura", "Factura"),
        ("FWD: Factura", "Factura"),
        ("Re : Factura", "Factura"),
        ("RE： Factura", "Factura"),
        ("  Re: Factura", "Factura"),
        ("RE: RV: Fwd: Factura", "Factura"),
        ("Re: Re: Re: Recibo", "Recibo"),
        ("RE： Fwd: Recibo", "Recibo"),
    ],
)
def test_prefix_with_colon_is_stripped(subject: str, expected: str) -> None:
    assert clean_subject(subject) == expected


def test_msg_suffix_strip_still_applies_after_prefix() -> None:
    assert clean_subject("Re: Recibo.msg", strip_msg_suffix=True) == "Recibo"
    assert clean_subject("Recibo.msg", strip_msg_suffix=True) == "Recibo"


@pytest.mark.parametrize("empty", [None, "", "   "])
def test_empty_subject_is_empty_string(empty: str | None) -> None:
    assert clean_subject(empty) == ""
