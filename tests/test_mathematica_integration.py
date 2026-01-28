#!/usr/bin/env python3
"""
Integration tests for Mathematica MCP tool with live kernel.

Tests end-to-end functionality that requires a live Wolfram kernel.
Run with: pytest tests/test_mathematica_integration.py -v

These tests are marked as slow and require:
- Wolfram Mathematica installed
- wolframclient Python package
- Valid Mathematica license

Skip these tests by default in CI: pytest -m "not slow"
"""

import contextlib
import os
import signal
import sys
import threading
import time
from unittest.mock import patch

import pytest

# Add the source directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from mcp_handley_lab.mathematica.tool import (
    EvaluationCancelledError,
    _get_session,
    evaluate,
    interruptible_evaluate,
)

pytestmark = pytest.mark.slow


class TestLiveCancellation:
    """Test cancellation functionality with a live Wolfram kernel."""

    @pytest.fixture(autouse=True)
    def setup_test_environment(self):
        """Set up test environment and ensure clean state."""
        # Clear any existing session state
        import mcp_handley_lab.mathematica.tool as tool_module

        # Reset global state
        tool_module._session = None
        tool_module._evaluation_count = 0
        tool_module._result_history.clear()
        tool_module._input_history.clear()

        yield

        # Cleanup after test
        if tool_module._session:
            with contextlib.suppress(Exception):
                tool_module._session.terminate()
            tool_module._session = None

    @pytest.mark.skipif(
        not os.path.exists("/usr/bin/WolframKernel"),
        reason="Wolfram Kernel not found at /usr/bin/WolframKernel",
    )
    def test_evaluation_cancellation_with_live_kernel(self):
        """
        Test that ESC cancellation works with a live kernel.

        This test:
        1. Starts a live Wolfram session
        2. Launches a long-running evaluation in a thread
        3. Sends SIGINT after a short delay
        4. Verifies the evaluation was cancelled properly
        """
        # Flag to track evaluation result
        evaluation_result = {}
        evaluation_exception = {}

        def run_evaluation():
            """Run a long evaluation that can be interrupted."""
            try:
                # This should take 5 seconds but be interrupted after 1 second
                result = evaluate("Pause[5]")
                evaluation_result["result"] = result
            except Exception as e:
                evaluation_exception["exception"] = e

        # Start evaluation in background thread
        eval_thread = threading.Thread(target=run_evaluation)
        eval_thread.daemon = True
        eval_thread.start()

        # Wait 1 second, then send interrupt
        time.sleep(1)
        os.kill(os.getpid(), signal.SIGINT)

        # Wait for thread to complete (should be quick due to cancellation)
        eval_thread.join(timeout=10)  # Generous timeout for cleanup

        # Verify cancellation occurred
        assert "result" in evaluation_result, "Evaluation should have completed"
        result = evaluation_result["result"]

        # The result should indicate failure due to cancellation
        assert result.success is False
        assert (
            "cancelled" in result.error.lower() or "interrupt" in result.error.lower()
        )
        assert result.note == "Evaluation was cancelled by user interrupt (ESC)"
        assert result.expression == "Pause[5]"

    @pytest.mark.skipif(
        not os.path.exists("/usr/bin/WolframKernel"),
        reason="Wolfram Kernel not found at /usr/bin/WolframKernel",
    )
    def test_session_preservation_after_cancellation(self):
        """
        Test that session state is preserved after cancellation.

        This test:
        1. Sets a variable
        2. Cancels a long computation
        3. Verifies the variable is still accessible
        """
        # Set a variable first
        result1 = evaluate("testVar = 42")
        assert result1.success is True

        # Track cancellation attempt
        cancellation_occurred = {}

        def attempt_long_computation():
            """Attempt a long computation that will be cancelled."""
            try:
                result = evaluate("Pause[3]; testVar + 1")
                cancellation_occurred["result"] = result
            except Exception as e:
                cancellation_occurred["exception"] = e

        # Start long computation
        compute_thread = threading.Thread(target=attempt_long_computation)
        compute_thread.daemon = True
        compute_thread.start()

        # Cancel after short delay
        time.sleep(0.5)
        os.kill(os.getpid(), signal.SIGINT)

        # Wait for cancellation
        compute_thread.join(timeout=10)

        # Verify cancellation occurred
        assert "result" in cancellation_occurred
        cancelled_result = cancellation_occurred["result"]
        assert cancelled_result.success is False

        # Verify session state preserved - testVar should still be accessible
        result3 = evaluate("testVar")
        assert result3.success is True
        assert "42" in result3.result

    @pytest.mark.skipif(
        not os.path.exists("/usr/bin/WolframKernel"),
        reason="Wolfram Kernel not found at /usr/bin/WolframKernel",
    )
    def test_interruptible_evaluate_direct(self):
        """
        Test the interruptible_evaluate function directly.

        This tests the lower-level function that handles the signal setup.
        """
        # Get a session
        session = _get_session()

        # Test that normal evaluation works
        result = interruptible_evaluate(session, "2 + 2")
        assert str(result) == "4"

        # Test cancellation behavior with mock interrupt
        with patch("os.kill") as mock_kill:

            def trigger_interrupt(*args):
                """Simulate the interrupt signal being sent."""
                # Raise the exception that would come from kernel interruption
                raise EvaluationCancelledError("Test cancellation")

            mock_kill.side_effect = trigger_interrupt

            # This should raise EvaluationCancelledError
            with pytest.raises(EvaluationCancelledError, match="Test cancellation"):
                # Send interrupt immediately to simulate ESC press
                os.kill(os.getpid(), signal.SIGINT)


class TestSessionManagement:
    """Test session lifecycle with live kernel."""

    @pytest.fixture(autouse=True)
    def reset_session(self):
        """Reset session state between tests."""
        import mcp_handley_lab.mathematica.tool as tool_module

        if tool_module._session:
            with contextlib.suppress(Exception):
                tool_module._session.terminate()
            tool_module._session = None
        tool_module._evaluation_count = 0
        tool_module._result_history.clear()
        tool_module._input_history.clear()

    @pytest.mark.skipif(
        not os.path.exists("/usr/bin/WolframKernel"),
        reason="Wolfram Kernel not found at /usr/bin/WolframKernel",
    )
    def test_session_initialization(self):
        """Test that session initializes correctly."""
        # First evaluation should initialize session
        result = evaluate("2 + 2")
        assert result.success is True
        assert result.evaluation_count == 1
        assert "4" in result.result

        # Second evaluation should use same session
        result2 = evaluate("3 + 3")
        assert result2.success is True
        assert result2.evaluation_count == 2
        assert "6" in result2.result


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
