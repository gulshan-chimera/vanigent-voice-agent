const WORD_DIGITS: Record<string, string> = {
  zero: "0", oh: "0", one: "1", two: "2", three: "3",
  four: "4", five: "5", six: "6", seven: "7", eight: "8", nine: "9",
};

export function toE164US(raw: string): string | null {
  if (!raw) return null;

  let t = raw.toLowerCase();

  // Spoken digits → figures. "eight hundred" → "800".
  t = t.replace(/\b(zero|oh|one|two|three|four|five|six|seven|eight|nine)\b/g, (w) => WORD_DIGITS[w]);
  t = t.replace(/(\d)\s*hundred\b/g, (_, d) => `${d}00`);

  // Strip everything that isn't a digit or a leading +.
  const plus = t.trim().startsWith("+");
  const digits = t.replace(/\D/g, "");

  // US numbers are 10 digits, or 11 with a leading 1.
  let national: string;
  if (digits.length === 10) {
    national = digits;
  } else if (digits.length === 11 && digits.startsWith("1")) {
    national = digits.slice(1);
  } else {
    return null; // Wrong length — can't be a US number.
  }

  const area = national.slice(0, 3);
  const exchange = national.slice(3, 6);

  // Area code and exchange must each start 2–9 (NANP rule). This rejects
  // IDs and malformed values that happen to be the right length.
  if (!/^[2-9]/.test(area) || !/^[2-9]/.test(exchange)) return null;

  const e164 = `+1${national}`;

  // Reject premium-rate and reserved ranges, so a bad number can't be
  // dialled at the caller's employer's expense.
  if (area === "900") return null;
  if (exchange === "976") return null;
  if (exchange === "555" && national.slice(6) >= "0100" && national.slice(6) <= "0199") return null;

  // The leading + in the input is informational only; we rebuilt from
  // digits, so `plus` isn't required — kept for clarity of intent.
  void plus;

  return e164;
}