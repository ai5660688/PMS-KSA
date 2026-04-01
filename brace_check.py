from pathlib import Path
text = Path('PMS/Controllers/HomeController.DailyWelding.cs').read_text()
stack = []
for idx, ch in enumerate(text):
    if ch == '{':
        stack.append(idx)
    elif ch == '}':
        if stack:
            stack.pop()
        else:
            print('extra closing brace at char', idx)
            break
else:
    if stack:
        print('unclosed braces count', len(stack))
        print('last positions', stack[-10:])
    else:
        print('braces balanced')
