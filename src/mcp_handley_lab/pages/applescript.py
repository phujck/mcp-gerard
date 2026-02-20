"""AppleScript integration for Pages operations."""

import subprocess
from pathlib import Path


class PagesAppleScript:
    """Execute AppleScript commands for Pages operations."""

    @staticmethod
    def run_script(script: str) -> str:
        """Run an AppleScript and return the result."""
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            raise RuntimeError(f"AppleScript error: {result.stderr.strip()}")
        return result.stdout.strip()

    @staticmethod
    def create_document(
        file_path: str,
        template: str = "Blank",
        initial_text: str = "",
    ) -> dict:
        """Create a new Pages document.

        Args:
            file_path: Path to save the new document
            template: Pages template name (default: "Blank")
            initial_text: Optional initial text content

        Returns:
            dict with success status and file_path
        """
        abs_path = str(Path(file_path).resolve())
        text_escaped = initial_text.replace("\\", "\\\\").replace('"', '\\"')

        # Create document first, then set text if needed
        if initial_text:
            script = f'''
tell application "Pages"
    set newDoc to make new document with properties {{document template:template "{template}"}}
    delay 0.5

    tell newDoc
        set selection to body text
        set body text to "{text_escaped}"
    end tell

    save newDoc in POSIX file "{abs_path}"
    close newDoc

    return "ok"
end tell
'''
        else:
            script = f'''
tell application "Pages"
    set newDoc to make new document with properties {{document template:template "{template}"}}
    save newDoc in POSIX file "{abs_path}"
    close newDoc
    return "ok"
end tell
'''
        PagesAppleScript.run_script(script)
        return {
            "success": True,
            "file_path": abs_path,
        }

    @staticmethod
    def get_body_text(file_path: str) -> dict:
        """Get the body text of a Pages document.

        Args:
            file_path: Path to the document

        Returns:
            dict with success status and text content
        """
        abs_path = str(Path(file_path).resolve())

        script = f'''
tell application "Pages"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        set bodyContent to body text
        close saving no
    end tell

    return bodyContent
end tell
'''
        text = PagesAppleScript.run_script(script)
        return {
            "success": True,
            "file_path": abs_path,
            "text": text,
        }

    @staticmethod
    def set_body_text(file_path: str, text: str) -> dict:
        """Replace the entire body text of a Pages document.

        Args:
            file_path: Path to the document
            text: New text content

        Returns:
            dict with success status
        """
        abs_path = str(Path(file_path).resolve())
        text_escaped = text.replace("\\", "\\\\").replace('"', '\\"')

        script = f'''
tell application "Pages"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        set body text to "{text_escaped}"
        save
        close
    end tell

    return "ok"
end tell
'''
        PagesAppleScript.run_script(script)
        return {
            "success": True,
            "file_path": abs_path,
        }

    @staticmethod
    def find_replace(
        file_path: str,
        find_text: str,
        replace_text: str,
    ) -> dict:
        """Find and replace text in a Pages document.

        Args:
            file_path: Path to the document
            find_text: Text to find
            replace_text: Text to replace with

        Returns:
            dict with success status and replacement count
        """
        abs_path = str(Path(file_path).resolve())
        find_escaped = find_text.replace("\\", "\\\\").replace('"', '\\"')
        replace_escaped = replace_text.replace("\\", "\\\\").replace('"', '\\"')

        # AppleScript to find and replace using string manipulation
        script = f'''
tell application "Pages"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        set bodyContent to body text as text
        set findStr to "{find_escaped}"
        set replaceStr to "{replace_escaped}"

        -- Count and replace using a handler
        set countNum to 0
        set newContent to ""
        set remainingText to bodyContent

        repeat
            set foundPos to offset of findStr in remainingText
            if foundPos = 0 then
                set newContent to newContent & remainingText
                exit repeat
            end if

            set countNum to countNum + 1

            -- Add text before the match
            if foundPos > 1 then
                set newContent to newContent & (text 1 thru (foundPos - 1) of remainingText)
            end if
            -- Add replacement
            set newContent to newContent & replaceStr

            -- Move past the match
            set afterPos to foundPos + (length of findStr)
            if afterPos > (length of remainingText) then
                exit repeat
            end if
            set remainingText to text afterPos thru -1 of remainingText
        end repeat

        if countNum > 0 then
            set body text to newContent
            save
        end if

        close
    end tell

    return countNum as text
end tell
'''
        count = int(PagesAppleScript.run_script(script))
        return {
            "success": True,
            "file_path": abs_path,
            "find_text": find_text,
            "replace_text": replace_text,
            "replacement_count": count,
        }

    @staticmethod
    def append_text(file_path: str, text: str) -> dict:
        """Append text to the end of a Pages document.

        Args:
            file_path: Path to the document
            text: Text to append

        Returns:
            dict with success status
        """
        abs_path = str(Path(file_path).resolve())
        text_escaped = text.replace("\\", "\\\\").replace('"', '\\"')

        script = f'''
tell application "Pages"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        set currentText to body text as text
        set body text to currentText & "{text_escaped}"
        save
        close
    end tell

    return "ok"
end tell
'''
        PagesAppleScript.run_script(script)
        return {
            "success": True,
            "file_path": abs_path,
        }

    @staticmethod
    def insert_text(file_path: str, text: str, position: int) -> dict:
        """Insert text at a specific character position.

        Args:
            file_path: Path to the document
            text: Text to insert
            position: Character position (1-indexed, 0 means end)

        Returns:
            dict with success status
        """
        abs_path = str(Path(file_path).resolve())
        text_escaped = text.replace("\\", "\\\\").replace('"', '\\"')

        if position == 0:
            # Append to end
            return PagesAppleScript.append_text(file_path, text)

        script = f'''
tell application "Pages"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        set currentText to body text as text
        set insertPos to {position}

        if insertPos > (length of currentText) then
            set newText to currentText & "{text_escaped}"
        else if insertPos <= 1 then
            set newText to "{text_escaped}" & currentText
        else
            set beforeText to text 1 thru (insertPos - 1) of currentText
            set afterText to text insertPos thru -1 of currentText
            set newText to beforeText & "{text_escaped}" & afterText
        end if

        set body text to newText
        save
        close
    end tell

    return "ok"
end tell
'''
        PagesAppleScript.run_script(script)
        return {
            "success": True,
            "file_path": abs_path,
            "position": position,
        }

    @staticmethod
    def get_word_count(file_path: str) -> dict:
        """Get word count and other statistics.

        Args:
            file_path: Path to the document

        Returns:
            dict with word count, character count, paragraph count
        """
        abs_path = str(Path(file_path).resolve())

        script = f'''
tell application "Pages"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        set wordCount to count of words of body text
        set charCount to count of characters of body text
        set paraCount to count of paragraphs of body text
        close saving no
    end tell

    return (wordCount as text) & "," & (charCount as text) & "," & (paraCount as text)
end tell
'''
        result = PagesAppleScript.run_script(script)
        words, chars, paras = result.split(",")
        return {
            "success": True,
            "file_path": abs_path,
            "word_count": int(words),
            "character_count": int(chars),
            "paragraph_count": int(paras),
        }

    @staticmethod
    def export_document(
        file_path: str,
        output_path: str,
        format: str = "PDF",
    ) -> dict:
        """Export a Pages document to another format.

        Args:
            file_path: Path to the Pages document
            output_path: Path for the exported file
            format: Export format (PDF, Word, plain text, EPUB)

        Returns:
            dict with success status and output path
        """
        abs_path = str(Path(file_path).resolve())
        abs_output = str(Path(output_path).resolve())

        # Map format names to Pages export format constants
        format_map = {
            "PDF": "PDF",
            "pdf": "PDF",
            "Word": "Microsoft Word",
            "word": "Microsoft Word",
            "docx": "Microsoft Word",
            "DOCX": "Microsoft Word",
            "text": "unformatted text",
            "txt": "unformatted text",
            "TEXT": "unformatted text",
            "EPUB": "EPUB",
            "epub": "EPUB",
        }

        export_format = format_map.get(format, "PDF")

        script = f'''
tell application "Pages"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        export to POSIX file "{abs_output}" as {export_format}
        close saving no
    end tell

    return "ok"
end tell
'''
        PagesAppleScript.run_script(script)
        return {
            "success": True,
            "input_path": abs_path,
            "output_path": abs_output,
            "format": format,
        }

    @staticmethod
    def get_document_info(file_path: str) -> dict:
        """Get document properties and information.

        Args:
            file_path: Path to the document

        Returns:
            dict with document name, modified status, etc.
        """
        abs_path = str(Path(file_path).resolve())

        script = f'''
tell application "Pages"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        set docName to name
        set isModified to modified
        set pageCount to count of pages
        close saving no
    end tell

    return docName & "|||" & (isModified as text) & "|||" & (pageCount as text)
end tell
'''
        result = PagesAppleScript.run_script(script)
        parts = result.split("|||")
        return {
            "success": True,
            "file_path": abs_path,
            "name": parts[0],
            "modified": parts[1] == "true",
            "page_count": int(parts[2]),
        }
