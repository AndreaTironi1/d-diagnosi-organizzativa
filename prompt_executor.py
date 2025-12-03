#!/usr/bin/env python3
"""
Simple prompt executor that substitutes variables into a prompt template.
"""

import argparse
import sys


def execute_prompt(template: str, variable: str) -> str:
    """
    Execute a prompt by substituting the variable into the template.

    Args:
        template: Prompt template with {variable} placeholder
        variable: Value to substitute into the template

    Returns:
        The executed prompt with variable substituted
    """
    try:
        result = template.format(variable=variable)
        return result
    except KeyError as e:
        print(f"Error: Template is missing placeholder {e}", file=sys.stderr)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="Execute a prompt with a variable substitution"
    )
    parser.add_argument(
        "prompt",
        help="Prompt template (use {variable} as placeholder)"
    )
    parser.add_argument(
        "variable",
        help="Variable value to substitute into the prompt"
    )
    parser.add_argument(
        "-o", "--output",
        help="Output file (default: stdout)",
        default=None
    )

    args = parser.parse_args()

    result = execute_prompt(args.prompt, args.variable)

    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(result)
        print(f"Result written to {args.output}")
    else:
        print(result)


if __name__ == "__main__":
    main()
