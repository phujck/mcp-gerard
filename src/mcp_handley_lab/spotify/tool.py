"""Spotify MCP server for playback control via AppleScript (macOS).

Controls the Spotify desktop app directly - no API credentials required.
Some features (search, playlist management) require the Web API and are
disabled until Spotify developer registration is available.

Requires: Spotify desktop app installed on macOS.
"""

import subprocess

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, Field

from mcp_handley_lab.shared.models import ServerInfo

mcp = FastMCP("Spotify")


# Pydantic models for structured responses


class Artist(BaseModel):
    """Spotify artist information."""

    name: str = Field(..., description="Artist name.")


class Album(BaseModel):
    """Spotify album information."""

    name: str = Field(..., description="Album name.")


class Track(BaseModel):
    """Spotify track information."""

    name: str = Field(..., description="Track name.")
    artist: str = Field(default="", description="Primary artist name.")
    album: str = Field(default="", description="Album name.")
    duration_ms: int = Field(default=0, description="Track duration in milliseconds.")
    position_ms: int = Field(
        default=0, description="Current playback position in milliseconds."
    )
    spotify_url: str = Field(default="", description="Spotify URL for the track.")
    artwork_url: str = Field(default="", description="Album artwork URL.")


class PlaybackState(BaseModel):
    """Current playback state."""

    is_playing: bool = Field(..., description="Whether playback is active.")
    is_running: bool = Field(..., description="Whether Spotify app is running.")
    track: Track | None = Field(default=None, description="Currently playing track.")
    shuffle: bool = Field(default=False, description="Whether shuffle is enabled.")
    repeat: str = Field(
        default="off", description="Repeat mode: off, track, or context."
    )
    volume: int = Field(default=0, description="Current volume (0-100).")


class OperationResult(BaseModel):
    """Result of a Spotify operation."""

    success: bool = Field(..., description="Whether the operation succeeded.")
    message: str = Field(..., description="Status message.")


# AppleScript helpers


def _run_applescript(script: str) -> str:
    """Run AppleScript and return output."""
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript error: {result.stderr.strip()}")
    return result.stdout.strip()


def _is_spotify_running() -> bool:
    """Check if Spotify is running."""
    script = """
    tell application "System Events"
        return (name of processes) contains "Spotify"
    end tell
    """
    result = _run_applescript(script)
    return result == "true"


def _ensure_spotify_running() -> bool:
    """Ensure Spotify is running, launch if needed."""
    if not _is_spotify_running():
        _run_applescript('tell application "Spotify" to activate')
        import time

        time.sleep(1)  # Give it a moment to start
    return _is_spotify_running()


# Playback Control Tools


@mcp.tool(description="Get the currently playing track and playback state.")
def now_playing() -> PlaybackState:
    """Get current playback state including track and playback options."""
    if not _is_spotify_running():
        return PlaybackState(is_playing=False, is_running=False)

    try:
        script = """
        tell application "Spotify"
            set isPlaying to player state is playing
            set shuffleState to shuffling
            set repeatState to repeating
            set vol to sound volume

            if player state is stopped then
                return "playing:" & isPlaying & "|shuffle:" & shuffleState & "|repeat:" & repeatState & "|volume:" & vol & "|track:none"
            else
                set trackName to name of current track
                set artistName to artist of current track
                set albumName to album of current track
                set trackDuration to duration of current track
                set trackPosition to player position
                set trackUrl to spotify url of current track
                set artworkUrl to artwork url of current track

                return "playing:" & isPlaying & "|shuffle:" & shuffleState & "|repeat:" & repeatState & "|volume:" & vol & "|name:" & trackName & "|artist:" & artistName & "|album:" & albumName & "|duration:" & trackDuration & "|position:" & trackPosition & "|url:" & trackUrl & "|artwork:" & artworkUrl
            end if
        end tell
        """
        output = _run_applescript(script)

        # Parse the output
        parts = dict(p.split(":", 1) for p in output.split("|") if ":" in p)

        is_playing = parts.get("playing", "false") == "true"
        shuffle = parts.get("shuffle", "false") == "true"
        repeat_raw = parts.get("repeat", "false")
        repeat = "context" if repeat_raw == "true" else "off"
        volume = int(float(parts.get("volume", "0")))

        track = None
        if parts.get("name") and parts.get("track") != "none":
            # Duration is already in ms from Spotify
            duration_ms = int(float(parts.get("duration", "0")))
            # Position is in seconds but may have locale-specific decimal (comma vs period)
            position_str = parts.get("position", "0").replace(",", ".")
            position_ms = int(float(position_str) * 1000)
            track = Track(
                name=parts.get("name", ""),
                artist=parts.get("artist", ""),
                album=parts.get("album", ""),
                duration_ms=duration_ms,
                position_ms=position_ms,
                spotify_url=parts.get("url", ""),
                artwork_url=parts.get("artwork", ""),
            )

        return PlaybackState(
            is_playing=is_playing,
            is_running=True,
            track=track,
            shuffle=shuffle,
            repeat=repeat,
            volume=volume,
        )
    except Exception as e:
        return PlaybackState(is_playing=False, is_running=True, track=None)


@mcp.tool(
    description="Start or resume playback. Optionally play a specific Spotify URI."
)
def play(
    uri: str = Field(
        default="",
        description="Spotify URI to play (e.g., spotify:track:xxx). "
        "Leave empty to resume current playback.",
    ),
) -> OperationResult:
    """Start or resume playback."""
    if not _ensure_spotify_running():
        return OperationResult(success=False, message="Could not launch Spotify")

    try:
        if uri:
            script = f'tell application "Spotify" to play track "{uri}"'
        else:
            script = 'tell application "Spotify" to play'
        _run_applescript(script)
        return OperationResult(
            success=True,
            message=f"Playback started{' for ' + uri if uri else ''}",
        )
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to start playback: {e}")


@mcp.tool(description="Pause playback.")
def pause() -> OperationResult:
    """Pause playback."""
    if not _is_spotify_running():
        return OperationResult(success=False, message="Spotify is not running")

    try:
        _run_applescript('tell application "Spotify" to pause')
        return OperationResult(success=True, message="Playback paused")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to pause: {e}")


@mcp.tool(description="Toggle play/pause.")
def playpause() -> OperationResult:
    """Toggle between play and pause."""
    if not _ensure_spotify_running():
        return OperationResult(success=False, message="Could not launch Spotify")

    try:
        _run_applescript('tell application "Spotify" to playpause')
        return OperationResult(success=True, message="Toggled play/pause")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to toggle: {e}")


@mcp.tool(description="Skip to the next track.")
def skip() -> OperationResult:
    """Skip to next track."""
    if not _is_spotify_running():
        return OperationResult(success=False, message="Spotify is not running")

    try:
        _run_applescript('tell application "Spotify" to next track')
        return OperationResult(success=True, message="Skipped to next track")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to skip: {e}")


@mcp.tool(description="Go back to the previous track.")
def previous() -> OperationResult:
    """Go to previous track."""
    if not _is_spotify_running():
        return OperationResult(success=False, message="Spotify is not running")

    try:
        _run_applescript('tell application "Spotify" to previous track')
        return OperationResult(success=True, message="Went to previous track")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to go back: {e}")


@mcp.tool(description="Set the playback volume (0-100).")
def volume(
    level: int = Field(
        ...,
        ge=0,
        le=100,
        description="Volume level from 0 to 100.",
    ),
) -> OperationResult:
    """Set playback volume."""
    if not _is_spotify_running():
        return OperationResult(success=False, message="Spotify is not running")

    try:
        _run_applescript(f'tell application "Spotify" to set sound volume to {level}')
        return OperationResult(success=True, message=f"Volume set to {level}%")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to set volume: {e}")


@mcp.tool(description="Adjust volume up or down by a relative amount.")
def volume_adjust(
    delta: int = Field(
        ...,
        ge=-100,
        le=100,
        description="Amount to adjust volume by (-100 to +100).",
    ),
) -> OperationResult:
    """Adjust volume relatively."""
    if not _is_spotify_running():
        return OperationResult(success=False, message="Spotify is not running")

    try:
        script = f"""
        tell application "Spotify"
            set currentVol to sound volume
            set newVol to currentVol + {delta}
            if newVol < 0 then set newVol to 0
            if newVol > 100 then set newVol to 100
            set sound volume to newVol
            return newVol
        end tell
        """
        new_vol = _run_applescript(script)
        return OperationResult(success=True, message=f"Volume adjusted to {new_vol}%")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to adjust volume: {e}")


@mcp.tool(description="Seek to a position in the current track.")
def seek(
    position_seconds: float = Field(
        ...,
        ge=0,
        description="Position in seconds to seek to.",
    ),
) -> OperationResult:
    """Seek to position in track."""
    if not _is_spotify_running():
        return OperationResult(success=False, message="Spotify is not running")

    try:
        _run_applescript(
            f'tell application "Spotify" to set player position to {position_seconds}'
        )
        mins, secs = divmod(int(position_seconds), 60)
        return OperationResult(success=True, message=f"Seeked to {mins}:{secs:02d}")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to seek: {e}")


@mcp.tool(description="Toggle shuffle mode.")
def shuffle(
    enabled: bool | None = Field(
        default=None,
        description="Set shuffle state. None to toggle.",
    ),
) -> OperationResult:
    """Toggle or set shuffle mode."""
    if not _is_spotify_running():
        return OperationResult(success=False, message="Spotify is not running")

    try:
        if enabled is None:
            script = """
            tell application "Spotify"
                set shuffling to not shuffling
                return shuffling
            end tell
            """
        else:
            script = f"""
            tell application "Spotify"
                set shuffling to {str(enabled).lower()}
                return shuffling
            end tell
            """
        result = _run_applescript(script)
        state = "on" if result == "true" else "off"
        return OperationResult(success=True, message=f"Shuffle is now {state}")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to set shuffle: {e}")


@mcp.tool(description="Toggle repeat mode.")
def repeat(
    enabled: bool | None = Field(
        default=None,
        description="Set repeat state. None to toggle.",
    ),
) -> OperationResult:
    """Toggle or set repeat mode."""
    if not _is_spotify_running():
        return OperationResult(success=False, message="Spotify is not running")

    try:
        if enabled is None:
            script = """
            tell application "Spotify"
                set repeating to not repeating
                return repeating
            end tell
            """
        else:
            script = f"""
            tell application "Spotify"
                set repeating to {str(enabled).lower()}
                return repeating
            end tell
            """
        result = _run_applescript(script)
        state = "on" if result == "true" else "off"
        return OperationResult(success=True, message=f"Repeat is now {state}")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to set repeat: {e}")


@mcp.tool(description="Open a Spotify URI or URL in the app.")
def open_uri(
    uri: str = Field(
        ...,
        description="Spotify URI (spotify:track:xxx) or URL (open.spotify.com/...) to open.",
    ),
) -> OperationResult:
    """Open a Spotify URI in the app."""
    if not _ensure_spotify_running():
        return OperationResult(success=False, message="Could not launch Spotify")

    try:
        # Convert URL to URI if needed
        if "open.spotify.com" in uri:
            # https://open.spotify.com/track/xxx -> spotify:track:xxx
            parts = uri.split("/")
            if len(parts) >= 2:
                item_type = parts[-2]  # track, album, playlist, artist
                item_id = parts[-1].split("?")[0]  # Remove query params
                uri = f"spotify:{item_type}:{item_id}"

        _run_applescript(f'open location "{uri}"')
        return OperationResult(success=True, message=f"Opened {uri}")
    except Exception as e:
        return OperationResult(success=False, message=f"Failed to open URI: {e}")


# Server Info


@mcp.tool(
    description="Get Spotify server status and available capabilities."
)
def server_info() -> ServerInfo:
    """Get server information and status."""
    is_running = _is_spotify_running()

    status = "active" if is_running else "spotify_not_running"
    dependencies = {
        "backend": "AppleScript (macOS)",
        "spotify_running": str(is_running),
        "api_features": "disabled (developer registration unavailable)",
    }

    return ServerInfo(
        name="Spotify",
        version="1.0.0",
        status=status,
        capabilities=[
            "now_playing",
            "play",
            "pause",
            "playpause",
            "skip",
            "previous",
            "volume",
            "volume_adjust",
            "seek",
            "shuffle",
            "repeat",
            "open_uri",
            "server_info",
        ],
        dependencies=dependencies,
    )
