#!/usr/bin/env python3
"""
Debug script to test Claude API connection and response format
"""

import os
import asyncio
import aiohttp
import json
import sys


async def test_claude_api():
    """Test Claude API with a simple request"""

    api_key = os.environ.get('CLAUDE_API_KEY')
    if not api_key:
        print("❌ CLAUDE_API_KEY environment variable not set!")
        print("Set it with: export CLAUDE_API_KEY='your-key-here'")
        return

    print(f"✅ API Key found: {api_key[:10]}...{api_key[-4:]}")

    # Test with a simple prompt first
    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01"
    }

    # Simple test payload
    test_payload = {
        "model": "claude-3-sonnet-20240229",
        "messages": [
            {
                "role": "user",
                "content": "Respond with only this JSON: {\"test\": \"success\", \"message\": \"API working\"}"
            }
        ],
        "max_tokens": 100
    }

    print("\n🔍 Testing Claude API...")
    print(f"URL: https://api.anthropic.com/v1/messages")
    print(f"Model: {test_payload['model']}")

    async with aiohttp.ClientSession() as session:
        try:
            async with session.post(
                    "https://api.anthropic.com/v1/messages",
                    headers=headers,
                    json=test_payload
            ) as response:
                response_text = await response.text()

                print(f"\n📡 Response Status: {response.status}")
                print(f"Response Headers: {dict(response.headers)}")

                if response.status == 200:
                    print("\n✅ API call successful!")
                    try:
                        data = json.loads(response_text)
                        print("\nResponse structure:")
                        print(json.dumps(data, indent=2))

                        # Extract the actual response
                        if 'content' in data and len(data['content']) > 0:
                            content = data['content'][0]['text']
                            print(f"\nClaude's response: {content}")

                            # Try to parse Claude's JSON response
                            try:
                                parsed_content = json.loads(content)
                                print(f"\nParsed JSON: {parsed_content}")
                            except json.JSONDecodeError:
                                print("\n⚠️  Claude didn't return valid JSON")

                    except json.JSONDecodeError as e:
                        print(f"\n❌ Failed to parse API response as JSON: {e}")
                        print(f"Raw response: {response_text[:500]}")

                else:
                    print(f"\n❌ API Error!")
                    print(f"Response text: {response_text[:1000]}")

                    # Check if it's HTML (authentication error)
                    if response_text.strip().startswith('<'):
                        print("\n⚠️  Received HTML response - likely an authentication error")
                        print("Please check your API key is correct")
                    else:
                        # Try to parse error
                        try:
                            error_data = json.loads(response_text)
                            if 'error' in error_data:
                                print(f"\nError details: {error_data['error']}")
                        except:
                            pass

        except aiohttp.ClientError as e:
            print(f"\n❌ Network error: {e}")
        except Exception as e:
            print(f"\n❌ Unexpected error: {e}")
            import traceback
            traceback.print_exc()


async def test_sop_prompt():
    """Test the actual SOP prompt"""

    api_key = os.environ.get('CLAUDE_API_KEY')
    if not api_key:
        print("\n❌ Skipping SOP prompt test - no API key")
        return

    print("\n\n🔍 Testing SOP Prompt...")

    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01"
    }

    # Test with sample PowerPoint content
    sample_pptx_content = """
Presentation: Sample Work Instruction.pptx
Total Slides: 3

--- Slide 1 ---
Title: Cassette Assembly Procedure
Content:
  • Work Instruction for Manufacturing Team
  • Version 1.0

--- Slide 2 ---
Title: Responsibilities
Content:
  • Technician: Performs assembly steps
  • Senior Technician: Reviews and approves work
  • Engineer: Updates procedures

--- Slide 3 ---
Title: Definitions
Content:
  • DFF: Direct Filter Flow
  • QC: Quality Control
  • SOP: Standard Operating Procedure
"""

    sop_prompt = """You MUST respond with a valid JSON object containing these exact fields:

{
    "title": "[Extract the main title/subject from the PowerPoint]",
    "objective": "[Extract or infer the main objective/purpose of this procedure]",
    "scope": "This SOP applies to all Cassette Manufacturing operations at the US Stoneham facility",
    "responsibilities": "[Extract specific roles mentioned]",
    "definitions": "[Extract any defined terms or 'N/A']",
    "mango_id": "",
    "pptx_title": "[Use the PowerPoint filename or main title]"
}"""

    full_prompt = f"{sop_prompt}\n\nPowerPoint Content:\n{sample_pptx_content}"

    payload = {
        "model": "claude-3-sonnet-20240229",
        "messages": [{"role": "user", "content": full_prompt}],
        "max_tokens": 1000
    }

    async with aiohttp.ClientSession() as session:
        try:
            async with session.post(
                    "https://api.anthropic.com/v1/messages",
                    headers=headers,
                    json=payload
            ) as response:
                if response.status == 200:
                    data = await response.json()
                    content = data['content'][0]['text']
                    print("\n✅ SOP prompt test successful!")
                    print(f"\nClaude's response:\n{content}")

                    # Try to parse as JSON
                    try:
                        # Clean response
                        cleaned = content.strip()
                        if cleaned.startswith('```json'):
                            cleaned = cleaned[7:]
                        if cleaned.startswith('```'):
                            cleaned = cleaned[3:]
                        if cleaned.endswith('```'):
                            cleaned = cleaned[:-3]

                        parsed = json.loads(cleaned.strip())
                        print("\n✅ Successfully parsed as JSON:")
                        print(json.dumps(parsed, indent=2))
                    except json.JSONDecodeError as e:
                        print(f"\n❌ Failed to parse response as JSON: {e}")
                else:
                    error_text = await response.text()
                    print(f"\n❌ API Error: {response.status}")
                    print(f"Details: {error_text[:500]}")

        except Exception as e:
            print(f"\n❌ Error: {e}")
            import traceback
            traceback.print_exc()


async def main():
    """Run all tests"""
    print("Claude API Debug Tool")
    print("=" * 50)

    # Test basic API connection
    await test_claude_api()

    # Test SOP prompt
    await test_sop_prompt()

    print("\n" + "=" * 50)
    print("Debug session complete!")


if __name__ == "__main__":
    # Handle Windows event loop
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    asyncio.run(main())