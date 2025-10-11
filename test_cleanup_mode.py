#!/usr/bin/env python3
"""
Test script for the new Cleanup Mode (` + 4) functionality.

This script tests the cleanup mode by simulating the workflow:
1. Load environment variables and API keys
2. Test the PROMPT_CLEANUP system prompt
3. Test cleanup with various types of messy text (terminal output, code, verbose text)
4. Verify the LLM fallback chain works
"""

import os
import sys
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Import the necessary functions from voice_shortcuts
# We'll import the AI response functions directly
import google.generativeai as genai
from anthropic import Anthropic
import openai

# Define the cleanup prompt (same as in voice_shortcuts.py)
PROMPT_CLEANUP = """
You are a text processing assistant that cleans and compacts text for efficient LLM consumption.

Your task is to take the provided text (which may contain terminal output, debug logs, code, data, or mixed content) and transform it into a clean, concise format suitable for feeding to another LLM.

RULES:
1. **Remove noise**: Strip out debug logs, stack traces, repetitive terminal output, timestamps, file paths that aren't essential
2. **Preserve code**: Keep code blocks intact and properly formatted
3. **Preserve data**: Keep structured data (JSON, tables, lists) but remove redundant entries
4. **Condense prose**: Summarize verbose explanations into key points
5. **Maintain context**: Ensure the cleaned text retains all essential information needed to understand the content
6. **No meta-commentary**: Do NOT add introductions like "Here's the cleaned text:" - just output the cleaned content directly
7. **Format for clarity**: Use markdown formatting (headers, lists, code blocks) to organize the output

If a voice instruction is provided, use it to guide what to focus on or what aspects to emphasize in the cleanup.

Output ONLY the cleaned and compacted text, ready to be pasted.
"""

# Test configurations (matching voice_shortcuts.py)
GEMINI_MODEL = "gemini-2.0-flash-exp"  # Updated to match voice_shortcuts.py
CLAUDE_MODEL = "claude-3-7-sonnet-20250219"
OPENAI_MODEL = "gpt-4o"
TEMP = 0.3

# Sample messy text to test cleanup
TEST_TEXT_1 = """
[2025-01-10 14:32:15] INFO: Starting server...
[2025-01-10 14:32:15] DEBUG: Loading configuration from /usr/local/etc/config.json
[2025-01-10 14:32:16] DEBUG: Configuration loaded successfully
[2025-01-10 14:32:16] INFO: Server listening on port 8080
[2025-01-10 14:32:17] DEBUG: Received connection from 192.168.1.100
[2025-01-10 14:32:17] DEBUG: Processing request: GET /api/users
[2025-01-10 14:32:17] DEBUG: Query executed in 45ms
[2025-01-10 14:32:17] INFO: Request completed successfully
[2025-01-10 14:32:18] DEBUG: Connection closed
[2025-01-10 14:32:20] ERROR: Database connection timeout
Traceback (most recent call last):
  File "/usr/local/lib/python3.11/site-packages/sqlalchemy/engine/base.py", line 1910, in _execute_context
    self.dialect.do_execute(
  File "/usr/local/lib/python3.11/site-packages/sqlalchemy/engine/default.py", line 736, in do_execute
    cursor.execute(statement, parameters)
psycopg2.OperationalError: FATAL:  connection timeout
[2025-01-10 14:32:21] INFO: Retrying database connection...
[2025-01-10 14:32:22] INFO: Database connection restored
"""

TEST_TEXT_2 = """
The SciSymbio system architecture is a multi-layered AI-powered platform designed for laboratory automation and documentation. The system consists of several key components that work together seamlessly. First, there's the user interface layer which includes both web and mobile applications built with modern frameworks like Next.js and Flutter. These interfaces provide scientists with intuitive tools for experiment planning, video review, and knowledge discovery. The web application features a protocol editor, reagent calculator, and template library for experiment planning. It also has a video review dashboard with timeline scrubber and AI annotations, plus object detection overlays for detailed analysis. The knowledge discovery hub works like "YouTube for experiments" - it's a searchable internal library of AI-annotated video summaries that helps accelerate training and transfer tacit knowledge between team members. Additionally, there's a knowledge graph explorer built with D3.js for interactive visualization, device management capabilities for hardware registration and firmware updates, and comprehensive analytics and reporting features. The mobile app handles device connectivity using both BLE and WiFi Direct, provides real-time voice interaction with wake word detection, and offers offline-first architecture so core features work without internet. It also displays live streaming previews and manages settings.
"""

def test_gemini_cleanup(text, instruction=""):
    """Test cleanup using Gemini API"""
    try:
        print("\n🧪 Testing with Gemini API...")

        # Check API key
        api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
        if not api_key:
            print("❌ GEMINI_API_KEY not found")
            return None

        genai.configure(api_key=api_key)

        # Build prompt
        prompt = f"{PROMPT_CLEANUP}\n\nText to clean:\n{text}"
        if instruction:
            prompt = f"{PROMPT_CLEANUP}\n\nVoice instruction: {instruction}\n\nText to clean:\n{text}"

        # Create model and generate
        model = genai.GenerativeModel(
            model_name=GEMINI_MODEL,
            generation_config={"temperature": TEMP, "max_output_tokens": 2048}
        )

        response = model.generate_content(prompt)

        if response.text:
            print("✅ Gemini cleanup successful")
            return response.text
        else:
            print("❌ Gemini returned empty response")
            return None

    except Exception as e:
        print(f"❌ Gemini error: {e}")
        return None

def test_claude_cleanup(text, instruction=""):
    """Test cleanup using Claude API"""
    try:
        print("\n🧪 Testing with Claude API...")

        # Check API key
        if not os.getenv("ANTHROPIC_API_KEY"):
            print("❌ ANTHROPIC_API_KEY not found")
            return None

        client = Anthropic()

        # Build message
        content = f"Text to clean:\n{text}"
        if instruction:
            content = f"Voice instruction: {instruction}\n\n{content}"

        messages = [{"role": "user", "content": [{"type": "text", "text": content}]}]

        # Generate response
        response = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=2048,
            temperature=TEMP,
            system=PROMPT_CLEANUP,
            messages=messages
        )

        if response.content[0].text:
            print("✅ Claude cleanup successful")
            return response.content[0].text
        else:
            print("❌ Claude returned empty response")
            return None

    except Exception as e:
        print(f"❌ Claude error: {e}")
        return None

def main():
    """Run all tests"""
    print("="*70)
    print("🧪 CLEANUP MODE TEST SUITE")
    print("="*70)

    # Check environment
    print("\n📋 Checking API Keys...")
    has_gemini = bool(os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY"))
    has_claude = bool(os.getenv("ANTHROPIC_API_KEY"))
    has_openai = bool(os.getenv("OPENAI_API_KEY"))

    print(f"  Gemini API Key: {'✅' if has_gemini else '❌'}")
    print(f"  Claude API Key: {'✅' if has_claude else '❌'}")
    print(f"  OpenAI API Key: {'✅' if has_openai else '❌'}")

    if not (has_gemini or has_claude or has_openai):
        print("\n❌ No API keys found! Please set up .env file")
        return

    # Test 1: Terminal logs cleanup
    print("\n" + "="*70)
    print("TEST 1: Cleaning Terminal Logs")
    print("="*70)
    print("\n📄 Original text (first 200 chars):")
    print(TEST_TEXT_1[:200] + "...")

    result = test_gemini_cleanup(TEST_TEXT_1, "focus on the error and key events")
    if result:
        print("\n✨ Cleaned result:")
        print("-"*70)
        print(result)
        print("-"*70)

    # Test 2: Verbose prose cleanup
    print("\n" + "="*70)
    print("TEST 2: Condensing Verbose Text")
    print("="*70)
    print("\n📄 Original text (first 200 chars):")
    print(TEST_TEXT_2[:200] + "...")

    result = test_gemini_cleanup(TEST_TEXT_2, "make it very concise, bullet points")
    if result:
        print("\n✨ Cleaned result:")
        print("-"*70)
        print(result)
        print("-"*70)

    # Test 3: Fallback chain (if Gemini is not available)
    print("\n" + "="*70)
    print("TEST 3: Testing Fallback Chain")
    print("="*70)

    if not has_gemini and has_claude:
        print("Testing Claude as fallback...")
        result = test_claude_cleanup(TEST_TEXT_1)
        if result:
            print("\n✅ Claude fallback works!")
    elif has_gemini:
        print("✅ Gemini is primary provider")

    print("\n" + "="*70)
    print("✅ TESTS COMPLETE")
    print("="*70)

    print("\n📝 To test in the actual app (Updated Workflow):")
    print("  1. Start voice_shortcuts.py")
    print("  2. Select and COPY text to clipboard (Cmd+C)")
    print("  3. Press ` + 4 (starts recording, shows '🎤 Speak...')")
    print("  4. Say your instruction: 'focus on errors' / 'make bullet points' / etc.")
    print("  5. Press ` + 4 again (stops recording)")
    print("  6. App processes and copies result to clipboard (shows '→ Clipboard')")
    print("  7. Paste result wherever you want (Cmd+V)")
    print("\n💡 Key difference from other modes:")
    print("  - Cleanup mode does NOT auto-paste")
    print("  - Voice instruction drives the transformation")
    print("  - Trigger keys (` + 4) stay in the field")

if __name__ == "__main__":
    main()
