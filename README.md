# AI Presentation Generator

A Python application that uses AI to generate professional PowerPoint presentations with structured slide content, detailed speaker notes, customizable themes, and optional images.

## Features

- **AI Content Generation:** Generates slide titles, bullet points, and speaker notes.
- **Automatic Theme Selection:** Chooses a theme based on the presentation topic.
- **Six Professional Themes:** Business, Tech, Creative, Nature, Elegant, and Healthcare.
- **Image Integration:** Retrieves relevant images from Unsplash, with placeholder images as a fallback.
- **PowerPoint Export:** Creates editable `.pptx` presentations automatically.

## Technologies Used

- Python
- Z.ai API (GLM-4.6)
- python-pptx
- Requests
- Pillow
- python-dotenv
- Unsplash API

## Getting Started

**1. Clone the repository**

```bash
git clone https://github.com/YOUR_USERNAME/Ai-presentation-Generator.git
cd Ai-presentation-Generator
```

**2. Install dependencies**

```bash
pip install -r requirements.txt
```

**3. Configure environment variables**

Create a `.env` file in the project root:

```env
ZAI_API_KEY=your_zai_api_key
UNSPLASH_ACCESS_KEY=your_unsplash_access_key
```

Keep your API keys private and never commit your `.env` file.

**4. Run the application**

```bash
python app.py
```

Follow the prompts to generate your presentation.

## How It Works

1. Enter the presentation title, description, presenter name, and number of slides.
2. The AI generates slide content and speaker notes.
3. The application applies a suitable theme and optionally retrieves images.
4. The final presentation is exported as a PowerPoint file.

## Future Improvements

- Develop a web interface.
- Add more slide layouts and customization options.
- Improve validation of AI-generated content.

## Author

**Ouissal Bouamar** — Full Stack Developer

[GitHub Profile](https://github.com/Ouissaal)

