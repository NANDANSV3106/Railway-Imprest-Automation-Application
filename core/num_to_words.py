"""Convert integers to Indian-English words (lakh/thousand system)."""


def num_to_words(n):
    """Convert integer to Indian English words (UPPERCASE)."""
    if n == 0:
        return "ZERO"
    ones = [
        '', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE',
        'TEN', 'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN',
        'SEVENTEEN', 'EIGHTEEN', 'NINETEEN'
    ]
    tens_w = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY',
              'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY']

    def below_hundred(x):
        if x < 20:
            return ones[x]
        return tens_w[x // 10] + (' ' + ones[x % 10] if x % 10 else '')

    def below_thousand(x):
        if x < 100:
            return below_hundred(x)
        return ones[x // 100] + ' HUNDRED' + (' ' + below_hundred(x % 100) if x % 100 else '')

    parts = []
    if n >= 100000:
        parts.append(below_thousand(n // 100000) + ' LAKH')
        n %= 100000
    if n >= 1000:
        parts.append(below_thousand(n // 1000) + ' THOUSAND')
        n %= 1000
    if n >= 100:
        parts.append(ones[n // 100] + ' HUNDRED')
        n %= 100
    if n:
        parts.append(below_hundred(n))
    return ' '.join(parts)


def amount_in_words(n):
    return f"({num_to_words(n)} ONLY)"


def amount_in_words_title(n):
    """Title-case version for labour annexure."""
    return f"({num_to_words(n).title()} only)"

