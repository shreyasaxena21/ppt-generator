# Auto-PPT-Generator

A publicly accessible web app that lets anyone turn bulk text or markdown into a fully formatted PowerPoint presentation based on an uploaded template.

## Features

- **Text to Slides:** Intelligently breaks down long-form text into a logical slide structure.
- **Template-Based Styling:** Infers and applies the style, fonts, and colors from an uploaded `.pptx` or `.potx` file.
- **Reusable Assets:** Reuses images found within the uploaded template.
- **LLM Agnostic:** Supports multiple LLM providers (OpenAI, Anthropic, Gemini) using the user's own API key.
- **Privacy-First:** Your API key and text content are never stored or logged on the server.

## How It Works

1.  **Input:** The user provides bulk text, their LLM API key, and a PowerPoint template.
2.  **LLM Call:** The app sends the text and a structured prompt to the chosen LLM. The prompt asks the LLM to format the content as a JSON array of slides, each with a title and bullet points.
3.  **Parsing & Structure:** The LLM's JSON response is parsed to define the presentation's structure.
4.  **Template Analysis:** The backend uses the `python-pptx` library to analyze the uploaded template. It identifies slide layouts, theme colors, and font styles. It also extracts any reusable images.
5.  **Presentation Generation:**
    -   A new `Presentation` object is created, starting with the uploaded template.
    -   Existing slides are cleared to ensure a clean slate.
    -   New slides are added based on the parsed LLM output.
    -   The application applies the inferred styles (fonts, colors, etc.) to the new slides' titles and text bodies.
    -   Images extracted from the template are added to the new slides where appropriate.
6.  **Download:** The newly generated `.pptx` file is sent back to the user for download.

## Setup and Usage

### Prerequisites

-   Python 3.7+
-   `pip`

### Local Installation

1.  **Clone the repository:**
    ```sh
    git clone [https://github.com/your-username/Auto-PPT-Generator.git](https://github.com/your-username/Auto-PPT-Generator.git)
    cd Auto-PPT-Generator
    ```

2.  **Install dependencies:**
    ```sh
    pip install -r requirements.txt
    ```
    (Note: Create a `requirements.txt` file with the following contents: `Flask`, `python-pptx`, `requests`)

3.  **Run the application:**
    ```sh
    python app.py
    ```

4.  **Access the app:**
    Open your web browser and navigate to `http://127.0.0.1:5000`.

### Hosted Demo

A working hosted link will be provided here for demonstration purposes.

*Disclaimer: This is a demo application. Use at your own risk. The app does not store any sensitive user data.*

## License

This project is licensed under the MIT License - see the `LICENSE` file for details.