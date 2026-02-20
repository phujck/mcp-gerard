"""AppleScript integration for Keynote operations."""

import subprocess
from pathlib import Path


class KeynoteAppleScript:
    """Execute AppleScript commands for Keynote operations."""

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
    def create_presentation(
        file_path: str,
        theme: str = "Basic White",
        title: str = "",
        subtitle: str = "",
    ) -> dict:
        """Create a new Keynote presentation.

        Args:
            file_path: Path to save the new presentation
            theme: Keynote theme name (default: "Basic White")
            title: Optional title for the first slide
            subtitle: Optional subtitle/body for the first slide

        Returns:
            dict with success status, file_path, and slide_count
        """
        abs_path = str(Path(file_path).resolve())

        # Escape strings for AppleScript
        title_escaped = title.replace('"', '\\"').replace("\n", "\\n")
        subtitle_escaped = subtitle.replace('"', '\\"').replace("\n", "\\n")

        script = f'''
tell application "Keynote"
    set newDoc to make new document with properties {{document theme:theme "{theme}"}}

    -- Set title slide content if provided
    tell slide 1 of newDoc
        if "{title_escaped}" is not "" then
            set object text of default title item to "{title_escaped}"
        end if
        if "{subtitle_escaped}" is not "" then
            set object text of default body item to "{subtitle_escaped}"
        end if
    end tell

    save newDoc in POSIX file "{abs_path}"
    set slideCount to count of slides of newDoc
    close newDoc

    return slideCount as text
end tell
'''
        slide_count = int(KeynoteAppleScript.run_script(script))
        return {
            "success": True,
            "file_path": abs_path,
            "slide_count": slide_count,
        }

    @staticmethod
    def add_slide(
        file_path: str,
        title: str = "",
        body: str = "",
        slide_layout: str = "Title & Bullets",
        position: int = 0,
    ) -> dict:
        """Add a new slide to an existing presentation.

        Args:
            file_path: Path to the presentation
            title: Slide title
            body: Slide body text
            slide_layout: Layout name (default: "Title & Bullets")
            position: Position to insert (0 = end, 1 = first, etc.)

        Returns:
            dict with success status, slide_number, and total slides
        """
        abs_path = str(Path(file_path).resolve())
        title_escaped = title.replace('"', '\\"').replace("\n", "\\n")
        body_escaped = body.replace('"', '\\"').replace("\n", "\\n")

        # Position logic: 0 means end, otherwise insert at position
        position_script = "at end" if position == 0 else f"at beginning"
        if position > 1:
            position_script = f"after slide {position - 1}"

        script = f'''
tell application "Keynote"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        set newSlide to make new slide {position_script} with properties {{base slide:master slide "{slide_layout}"}}

        tell newSlide
            if "{title_escaped}" is not "" then
                set object text of default title item to "{title_escaped}"
            end if
            if "{body_escaped}" is not "" then
                set object text of default body item to "{body_escaped}"
            end if
        end tell

        set slideNum to slide number of newSlide
        set totalSlides to count of slides
        save
        close
    end tell

    return (slideNum as text) & "," & (totalSlides as text)
end tell
'''
        result = KeynoteAppleScript.run_script(script)
        slide_num, total = result.split(",")
        return {
            "success": True,
            "file_path": abs_path,
            "slide_number": int(slide_num),
            "total_slides": int(total),
        }

    @staticmethod
    def set_slide_content(
        file_path: str,
        slide_number: int,
        title: str | None = None,
        body: str | None = None,
        notes: str | None = None,
    ) -> dict:
        """Set content of an existing slide.

        Args:
            file_path: Path to the presentation
            slide_number: Slide number (1-indexed)
            title: New title (None to keep existing)
            body: New body text (None to keep existing)
            notes: New presenter notes (None to keep existing)

        Returns:
            dict with success status
        """
        abs_path = str(Path(file_path).resolve())

        # Build the content-setting part of the script
        content_parts = []
        if title is not None:
            title_escaped = title.replace('"', '\\"').replace("\n", "\\n")
            content_parts.append(
                f'set object text of default title item to "{title_escaped}"'
            )
        if body is not None:
            body_escaped = body.replace('"', '\\"').replace("\n", "\\n")
            content_parts.append(
                f'set object text of default body item to "{body_escaped}"'
            )
        if notes is not None:
            notes_escaped = notes.replace('"', '\\"').replace("\n", "\\n")
            content_parts.append(f'set presenter notes to "{notes_escaped}"')

        if not content_parts:
            return {"success": True, "file_path": abs_path, "message": "No changes made"}

        content_script = "\n            ".join(content_parts)

        script = f'''
tell application "Keynote"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        tell slide {slide_number}
            {content_script}
        end tell
        save
        close
    end tell

    return "ok"
end tell
'''
        KeynoteAppleScript.run_script(script)
        return {
            "success": True,
            "file_path": abs_path,
            "slide_number": slide_number,
        }

    @staticmethod
    def delete_slide(file_path: str, slide_number: int) -> dict:
        """Delete a slide from the presentation.

        Args:
            file_path: Path to the presentation
            slide_number: Slide number to delete (1-indexed)

        Returns:
            dict with success status and remaining slide count
        """
        abs_path = str(Path(file_path).resolve())

        script = f'''
tell application "Keynote"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        delete slide {slide_number}
        set remainingSlides to count of slides
        save
        close
    end tell

    return remainingSlides as text
end tell
'''
        remaining = int(KeynoteAppleScript.run_script(script))
        return {
            "success": True,
            "file_path": abs_path,
            "remaining_slides": remaining,
        }

    @staticmethod
    def duplicate_slide(file_path: str, slide_number: int) -> dict:
        """Duplicate a slide in the presentation.

        Args:
            file_path: Path to the presentation
            slide_number: Slide number to duplicate (1-indexed)

        Returns:
            dict with new slide number and total slides
        """
        abs_path = str(Path(file_path).resolve())

        script = f'''
tell application "Keynote"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        duplicate slide {slide_number}
        set totalSlides to count of slides
        save
        close
    end tell

    return totalSlides as text
end tell
'''
        total = int(KeynoteAppleScript.run_script(script))
        return {
            "success": True,
            "file_path": abs_path,
            "new_slide_number": slide_number + 1,
            "total_slides": total,
        }

    @staticmethod
    def get_slide_layouts(file_path: str) -> dict:
        """Get available slide layouts (master slides) in a presentation.

        Args:
            file_path: Path to the presentation

        Returns:
            dict with list of layout names
        """
        abs_path = str(Path(file_path).resolve())

        script = f'''
tell application "Keynote"
    open POSIX file "{abs_path}"
    delay 0.5

    tell document 1
        set layoutNames to name of every master slide
        close saving no
    end tell

    set AppleScript's text item delimiters to "|||"
    return layoutNames as text
end tell
'''
        result = KeynoteAppleScript.run_script(script)
        layouts = result.split("|||") if result else []
        return {
            "success": True,
            "file_path": abs_path,
            "layouts": layouts,
        }
