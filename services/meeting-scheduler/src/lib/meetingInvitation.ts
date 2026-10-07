/** Format caller-supplied text only. Microsoft appends its own Teams meeting block. */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
const typography =
  "font-family:Segoe UI,Arial,sans-serif;font-size:12pt;line-height:1.5;color:#242424;";
function topic(text: string): string {
  const parts = /^(.*?)\s+([??-])\s+(.+)$/.exec(text);
  return parts
    ? "<strong>" +
        escapeHtml(parts[1]) +
        "</strong> " +
        parts[2] +
        " " +
        escapeHtml(parts[3])
    : escapeHtml(text);
}

export function invitationBody(description?: string): {
  contentType: "HTML";
  content: string;
} {
  const lines = (description ?? "")
    .replace(/\r\n|\r/g, "\n")
    .trim()
    .split("\n");
  const numbered = (line: string) => /^\s*(\d{1,3})[.)]\s+(.+)$/.exec(line);
  const blocks: string[] = [];
  let index = 0;
  while (index < lines.length) {
    if (!lines[index].trim()) {
      index++;
      continue;
    }
    const first = numbered(lines[index]);
    if (first) {
      const items: string[] = [];
      while (index < lines.length) {
        if (!lines[index].trim()) {
          index++;
          continue;
        }
        const item = numbered(lines[index]);
        if (!item) break;
        items.push(
          '<li value="' +
            Number(item[1]) +
            '" style="' +
            typography +
            'margin:0 0 12px 0;padding-left:2px;">' +
            topic(item[2]) +
            "</li>",
        );
        index++;
      }
      blocks.push(
        '<ol start="' +
          Number(first[1]) +
          '" style="' +
          typography +
          'margin:0 0 16px 0;padding-left:40px;">' +
          items.join("") +
          "</ol>",
      );
    } else {
      const paragraph: string[] = [];
      while (
        index < lines.length &&
        lines[index].trim() &&
        !numbered(lines[index])
      )
        paragraph.push(lines[index++]);
      let next = index;
      while (next < lines.length && !lines[next].trim()) next++;
      const heading =
        blocks.length === 0 &&
        paragraph.length === 1 &&
        next < lines.length &&
        !!numbered(lines[next]);
      const text = paragraph.map(escapeHtml).join("<br>");
      blocks.push(
        '<p style="' +
          typography +
          'margin:0 0 16px 0;">' +
          (heading ? "<strong>" + text + "</strong>" : text) +
          "</p>",
      );
    }
  }
  return {
    contentType: "HTML",
    content:
      '<html><body><div style="' +
      typography +
      '">' +
      (blocks.length
        ? '<div style="margin:0 0 40px 0;">' + blocks.join("") + "</div>"
        : "") +
      "</div></body></html>",
  };
}
