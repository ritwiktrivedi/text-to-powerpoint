import streamlit as st
import openai
import anthropic
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import json
import io
import zipfile
import xml.etree.ElementTree as ET
import re
from typing import Dict, List, Any
import base64
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="📊 Text to PowerPoint Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .feature-box {
        background: #f8fafc;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #4f46e5;
        margin: 1rem 0;
    }
    .warning-box {
        background: #fef3c7;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #f59e0b;
        margin: 1rem 0;
    }
    .success-box {
        background: #f0fdf4;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #10b981;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="main-header">
    <h1>📊 Text to PowerPoint Generator</h1>
    <p>Transform your text into beautiful presentations using AI and custom templates</p>
</div>
""", unsafe_allow_html=True)


class PresentationGenerator:
    def __init__(self):
        self.template_prs = None
        self.template_styles = {}

    def extract_template_styles(self, template_prs: Presentation) -> Dict[str, Any]:
        """Extract styles from the template, mapping layout and placeholder types to style info."""
        styles = {
            'slide_width': getattr(template_prs, 'slide_width', 9144000),
            'slide_height': getattr(template_prs, 'slide_height', 6858000),
            'layouts': [],
            'layout_map': {},  # name -> index
            # (layout_name, placeholder_type) -> font_info
            'placeholder_styles': {},
            'theme_colors': {},
            'background_fills': []
        }
        try:
            if hasattr(template_prs, 'slide_layouts'):
                for i, layout in enumerate(template_prs.slide_layouts):
                    layout_name = getattr(layout, 'name', f'Layout {i+1}')
                    styles['layout_map'][layout_name] = i
                    layout_info = {
                        'index': i,
                        'name': layout_name,
                        'placeholders': [],
                        'background': None
                    }
                    # Extract background fill if present
                    try:
                        if hasattr(layout, 'background') and layout.background.fill:
                            fill = layout.background.fill
                            if hasattr(fill, 'solid') and fill.solid:
                                color = fill.fore_color
                                if hasattr(color, 'rgb'):
                                    layout_info['background'] = {
                                        'type': 'solid',
                                        'color': str(color.rgb)
                                    }
                    except:
                        pass
                    # Extract placeholder styles
                    try:
                        for placeholder in layout.placeholders:
                            ph_type = str(
                                getattr(placeholder.placeholder_format, 'type', 'unknown'))
                            placeholder_info = {
                                'idx': getattr(placeholder.placeholder_format, 'idx', 0),
                                'type': ph_type,
                                'left': getattr(placeholder, 'left', 0),
                                'top': getattr(placeholder, 'top', 0),
                                'width': getattr(placeholder, 'width', 0),
                                'height': getattr(placeholder, 'height', 0),
                                'font_info': {},
                                'fill_info': {}
                            }
                            # Extract font styling from all paragraphs/runs if present
                            try:
                                if hasattr(placeholder, 'text_frame') and placeholder.text_frame:
                                    for para in placeholder.text_frame.paragraphs:
                                        for run in para.runs:
                                            if hasattr(run, 'font'):
                                                font = run.font
                                                font_info = {
                                                    'name': getattr(font, 'name', None),
                                                    'size': getattr(font, 'size', None),
                                                    'bold': getattr(font, 'bold', None),
                                                    'italic': getattr(font, 'italic', None),
                                                    'color': str(getattr(font.color, 'rgb', '')) if hasattr(font, 'color') and hasattr(font.color, 'rgb') else None
                                                }
                                                # Store by (layout_name, placeholder_type)
                                                styles['placeholder_styles'][(
                                                    layout_name, ph_type)] = font_info
                                                placeholder_info['font_info'] = font_info
                                                break
                                        if placeholder_info['font_info']:
                                            break
                            except:
                                pass
                            layout_info['placeholders'].append(
                                placeholder_info)
                    except Exception as e:
                        pass
                    styles['layouts'].append(layout_info)
            # Extract theme colors from slide master
            try:
                if hasattr(template_prs, 'slide_master'):
                    master = template_prs.slide_master
                    if hasattr(master, 'theme') and master.theme:
                        theme = master.theme
                        if hasattr(theme, 'color_scheme'):
                            color_scheme = theme.color_scheme
                            styles['theme_colors'] = {
                                'accent1': getattr(color_scheme, 'accent1_color', None),
                                'accent2': getattr(color_scheme, 'accent2_color', None),
                                'accent3': getattr(color_scheme, 'accent3_color', None),
                                'dark1': getattr(color_scheme, 'dk1_color', None),
                                'dark2': getattr(color_scheme, 'dk2_color', None),
                                'light1': getattr(color_scheme, 'lt1_color', None),
                                'light2': getattr(color_scheme, 'lt2_color', None),
                            }
                    # Extract master slide background
                    try:
                        if hasattr(master, 'background') and master.background.fill:
                            fill = master.background.fill
                            if hasattr(fill, 'solid') and fill.solid:
                                color = fill.fore_color
                                if hasattr(color, 'rgb'):
                                    styles['master_background'] = {
                                        'type': 'solid',
                                        'color': str(color.rgb)
                                    }
                    except:
                        pass
            except Exception as e:
                pass
        except Exception as e:
            # Provide fallback values
            styles['layouts'] = [
                {'index': 0, 'name': 'Title Slide', 'placeholders': []},
                {'index': 1, 'name': 'Content Slide', 'placeholders': []}
            ]
        return styles

    def call_ai_api(self, provider: str, api_key: str, prompt: str) -> str:
        """Call the appropriate AI API"""
        try:
            if provider == "OpenAI":
                client = openai.OpenAI(api_key=api_key)
                response = client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=4000,
                    temperature=0.7
                )
                return response.choices[0].message.content

            elif provider == "Anthropic":
                client = anthropic.Anthropic(api_key=api_key)
                message = client.messages.create(
                    model="claude-3-sonnet-20240229",
                    max_tokens=4000,
                    messages=[{"role": "user", "content": prompt}]
                )
                return message.content[0].text

            elif provider == "Google Gemini":
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-2.0-flash')
                response = model.generate_content(
                    prompt,
                    generation_config=genai.types.GenerationConfig(
                        temperature=0.7,
                        max_output_tokens=4000,
                    )
                )
                return response.text

        except Exception as e:
            raise Exception(f"AI API Error: {str(e)}")

    def create_prompt(self, input_text: str, guidance: str = "") -> str:
        """Create the prompt for AI processing"""
        return f"""Please analyze the following text and convert it into a structured PowerPoint presentation outline.

IMPORTANT: Respond with ONLY a valid JSON object in this exact format:
{{
  "title": "Presentation Title",
  "slides": [
    {{
      "title": "Slide Title",
      "content": ["bullet point 1", "bullet point 2", "bullet point 3"],
      "notes": "Speaker notes for this slide (optional)"
    }}
  ]
}}

Guidelines:
- Create 5-12 slides based on content length and complexity
- Each slide should have 2-5 concise bullet points
- Make titles engaging and descriptive
- Include speaker notes when helpful
- Focus on key insights, not just copying text
{f"- Style/tone guidance: {guidance}" if guidance else ""}

Text to convert:
{input_text}

Remember: Respond with ONLY the JSON object, no additional text or formatting."""

    def parse_ai_response(self, response: str) -> Dict[str, Any]:
        """Parse and validate AI response"""
        try:
            # Find JSON in the response
            json_start = response.find('{')
            json_end = response.rfind('}')

            if json_start == -1 or json_end == -1:
                raise ValueError("No JSON found in response")

            json_str = response[json_start:json_end + 1]
            structure = json.loads(json_str)

            if not structure.get('slides') or not isinstance(structure['slides'], list):
                raise ValueError("Invalid presentation structure")

            return structure

        except Exception as e:
            raise Exception(f"Failed to parse AI response: {str(e)}")

    def create_presentation(self, structure: Dict[str, Any], template_prs: Presentation = None) -> Presentation:
        """Create PowerPoint presentation from structure with proper styling"""
        if template_prs:
            # Create a new presentation but copy the template file directly to preserve all styling
            try:
                # Method 1: Try to clone the entire template presentation
                temp_buffer = io.BytesIO()
                template_prs.save(temp_buffer)
                temp_buffer.seek(0)
                prs = Presentation(temp_buffer)

                # Clear existing content slides (keep layouts and masters)
                slides_to_remove = []
                for i, slide in enumerate(prs.slides):
                    slides_to_remove.append(i)

                # Remove slides in reverse order to maintain indices
                for i in reversed(slides_to_remove):
                    r_id = prs.slides._sldIdLst[i].rId
                    prs.part.drop_rel(r_id)
                    del prs.slides._sldIdLst[i]

            except Exception as e:
                # Method 2: Fallback - create new presentation with template layouts
                st.warning(
                    f"Could not clone template completely, using fallback method: {str(e)}")
                prs = Presentation()
                try:
                    # Try to copy slide master from template
                    prs.slide_width = template_prs.slide_width
                    prs.slide_height = template_prs.slide_height
                except:
                    pass

            # Use template layouts
            try:
                layouts_to_use = template_prs.slide_layouts
            except:
                layouts_to_use = prs.slide_layouts
        else:
            # Create new presentation with default template
            prs = Presentation()
            layouts_to_use = prs.slide_layouts

        # Get template styles for applying formatting
        template_styles = getattr(self, 'template_styles', {})

        # Ensure we have at least 2 layouts (title and content)
        title_layout = layouts_to_use[0] if len(
            layouts_to_use) > 0 else prs.slide_layouts[0]
        content_layout = layouts_to_use[1] if len(
            layouts_to_use) > 1 else prs.slide_layouts[1]

        # Add title slide
        title_slide = prs.slides.add_slide(title_layout)
        layout_name = getattr(title_layout, 'name', None)
        # Set title with styling
        if title_slide.shapes.title:
            title_slide.shapes.title.text = structure.get(
                'title', 'Generated Presentation')
            # Apply template font styling to title
            for para in title_slide.shapes.title.text_frame.paragraphs:
                self.apply_paragraph_styling(
                    para, template_styles, 'title', layout_name=layout_name, placeholder_type='TITLE')
        # Add subtitle if available
        try:
            if len(title_slide.placeholders) > 1:
                subtitle_placeholder = title_slide.placeholders[1]
                if hasattr(subtitle_placeholder, 'text'):
                    subtitle_placeholder.text = f"Generated on {datetime.now().strftime('%B %d, %Y')}"
                    for para in subtitle_placeholder.text_frame.paragraphs:
                        self.apply_paragraph_styling(
                            para, template_styles, 'subtitle', layout_name=layout_name, placeholder_type='SUBTITLE')
        except:
            pass

        # Add content slides
        for slide_data in structure['slides']:
            slide = prs.slides.add_slide(content_layout)
            layout_name = getattr(content_layout, 'name', None)
            # Set slide title with styling
            if slide.shapes.title:
                slide.shapes.title.text = slide_data['title']
                for para in slide.shapes.title.text_frame.paragraphs:
                    self.apply_paragraph_styling(
                        para, template_styles, 'slide_title', layout_name=layout_name, placeholder_type='TITLE')
            # Add content to the appropriate placeholder
            content_added = False
            # Try to find content placeholder
            for placeholder in slide.placeholders:
                try:
                    ph_type = str(
                        getattr(placeholder.placeholder_format, 'type', 'unknown'))
                    if (placeholder.placeholder_format.idx == 1 or
                        'content' in ph_type.lower() or
                            'body' in ph_type.lower()):
                        text_frame = placeholder.text_frame
                        text_frame.clear()
                        for i, point in enumerate(slide_data['content']):
                            if i == 0:
                                p = text_frame.paragraphs[0]
                            else:
                                p = text_frame.add_paragraph()
                            p.text = point
                            p.level = 0
                            self.apply_paragraph_styling(
                                p, template_styles, 'content', layout_name=layout_name, placeholder_type=ph_type)
                        content_added = True
                        break
                except Exception as e:
                    continue
            # Fallback: add text box if no suitable placeholder found
            if not content_added:
                try:
                    left = Inches(1)
                    top = Inches(1.5)
                    width = Inches(8)
                    height = Inches(5)
                    textbox = slide.shapes.add_textbox(
                        left, top, width, height)
                    text_frame = textbox.text_frame
                    for i, point in enumerate(slide_data['content']):
                        if i == 0:
                            p = text_frame.paragraphs[0]
                        else:
                            p = text_frame.add_paragraph()
                        p.text = f"• {point}"
                        p.level = 0
                        self.apply_paragraph_styling(
                            p, template_styles, 'content', layout_name=layout_name, placeholder_type='BODY')
                except:
                    pass
            # Add speaker notes
            try:
                if slide_data.get('notes'):
                    notes_slide = slide.notes_slide
                    if hasattr(notes_slide, 'notes_text_frame'):
                        notes_slide.notes_text_frame.text = slide_data['notes']
            except:
                pass

        return prs

    def apply_text_styling(self, shape, template_styles: Dict, style_type: str):
        """Apply text styling from template to a text shape"""
        if not template_styles or not hasattr(shape, 'text_frame'):
            return

        try:
            text_frame = shape.text_frame
            if not text_frame.paragraphs:
                return

            for paragraph in text_frame.paragraphs:
                self.apply_paragraph_styling(
                    paragraph, template_styles, style_type)

        except Exception as e:
            pass  # Silently fail if styling can't be applied

    def apply_paragraph_styling(self, paragraph, template_styles: Dict, style_type: str, layout_name: str = None, placeholder_type: str = None):
        """Apply paragraph-level styling from template, using layout and placeholder type if available."""
        if not template_styles or not hasattr(paragraph, 'runs'):
            return
        try:
            font_info = None
            # Try to get font info by layout and placeholder type
            if layout_name and placeholder_type:
                font_info = template_styles.get('placeholder_styles', {}).get(
                    (layout_name, placeholder_type))
            # Fallback: use any available font info
            if not font_info:
                layouts = template_styles.get('layouts', [])
                for layout in layouts:
                    for placeholder in layout.get('placeholders', []):
                        if placeholder.get('font_info') and any(placeholder['font_info'].values()):
                            font_info = placeholder['font_info']
                            break
                    if font_info:
                        break
            if not font_info:
                return
            for run in paragraph.runs:
                if hasattr(run, 'font'):
                    font = run.font
                    if font_info.get('name'):
                        try:
                            font.name = font_info['name']
                        except:
                            fallback_fonts = [
                                'Calibri', 'Arial', 'Times New Roman', 'Helvetica']
                            for fallback in fallback_fonts:
                                try:
                                    font.name = fallback
                                    break
                                except:
                                    continue
                    if font_info.get('size') and font_info['size']:
                        try:
                            font.size = font_info['size']
                        except:
                            pass
                    if font_info.get('bold') is not None:
                        try:
                            font.bold = font_info['bold']
                        except:
                            pass
                    if font_info.get('italic') is not None:
                        try:
                            font.italic = font_info['italic']
                        except:
                            pass
                    if font_info.get('color') and font_info['color']:
                        try:
                            color_str = font_info['color'].replace(
                                'RGBColor(0x', '').replace(')', '')
                            if len(color_str) == 6:
                                r = int(color_str[0:2], 16)
                                g = int(color_str[2:4], 16)
                                b = int(color_str[4:6], 16)
                                font.color.rgb = RGBColor(r, g, b)
                        except:
                            pass
        except Exception as e:
            pass

    def test_presentation_creation(self) -> bool:
        """Test if we can create a basic presentation"""
        try:
            test_structure = {
                'title': 'Test Presentation',
                'slides': [
                    {
                        'title': 'Test Slide 1',
                        'content': ['Test point 1', 'Test point 2'],
                        'notes': 'Test notes'
                    }
                ]
            }

            prs = self.create_presentation(test_structure, self.template_prs)

            # Test saving
            output = io.BytesIO()
            prs.save(output)
            output.seek(0)

            return len(output.getvalue()) > 0
        except Exception as e:
            st.error(f"Presentation creation test failed: {str(e)}")
            return False


def main():
    generator = PresentationGenerator()

    # Sidebar for configuration
    with st.sidebar:
        st.header("🔧 Configuration")

        # AI Provider Selection
        st.subheader("AI Provider")
        provider = st.selectbox(
            "Choose AI Provider",
            ["OpenAI", "Anthropic", "Google Gemini"],
            help="Select your preferred AI service"
        )

        # API Key Input
        api_key = st.text_input(
            f"{provider} API Key",
            type="password",
            help="Your API key is never stored or logged"
        )

        if not api_key:
            st.markdown("""
            <div class="warning-box">
                ⚠️ <strong>API Key Required</strong><br>
                Get your API key from:
                <ul>
                    <li><strong>OpenAI:</strong> platform.openai.com</li>
                    <li><strong>Anthropic:</strong> console.anthropic.com</li>
                    <li><strong>Google:</strong> console.cloud.google.com</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

        st.divider()

        # Template Upload
        st.subheader("🎨 Template")
        template_file = st.file_uploader(
            "Upload PowerPoint Template",
            type=['pptx', 'potx'],
            help="Upload your branded PowerPoint template to preserve styling"
        )

        if template_file:
            try:
                # Load template
                generator.template_prs = Presentation(template_file)
                generator.template_styles = generator.extract_template_styles(
                    generator.template_prs)
                st.success("✅ Template loaded successfully!")

                # Show template info
                with st.expander("Template Information"):
                    st.write(
                        f"**Slide Size:** {generator.template_styles['slide_width']} x {generator.template_styles['slide_height']}")
                    st.write(
                        f"**Available Layouts:** {len(generator.template_styles['layouts'])}")

                    for i, layout in enumerate(generator.template_styles['layouts']):
                        st.write(
                            f"  - Layout {i+1}: {layout['name']} ({len(layout['placeholders'])} placeholders)")

                        # Show extracted font information
                        fonts_found = []
                        for placeholder in layout['placeholders']:
                            font_info = placeholder.get('font_info', {})
                            if any(font_info.values()):
                                font_details = []
                                if font_info.get('name'):
                                    font_details.append(
                                        f"Font: {font_info['name']}")
                                if font_info.get('size'):
                                    font_details.append(
                                        f"Size: {font_info['size']}")
                                if font_info.get('color'):
                                    font_details.append(
                                        f"Color: {font_info['color']}")
                                if font_details:
                                    fonts_found.append(
                                        f"    • {', '.join(font_details)}")

                        if fonts_found:
                            st.write("    Font Styles Found:")
                            # Show max 2 font styles per layout
                            for font in fonts_found[:2]:
                                st.write(font)

                    # Show theme colors if extracted
                    theme_colors = generator.template_styles.get(
                        'theme_colors', {})
                    if any(theme_colors.values()):
                        st.write("**Theme Colors:** Extracted ✅")
                        color_count = len(
                            [c for c in theme_colors.values() if c])
                        st.write(f"  - {color_count} theme colors available")
                    else:
                        st.write("**Theme Colors:** Could not extract")

                    # Show background info
                    if generator.template_styles.get('master_background'):
                        st.write("**Master Background:** Detected ✅")
                    else:
                        st.write("**Master Background:** Using default")

            except Exception as e:
                st.error(f"❌ Error loading template: {str(e)}")
                st.info(
                    "💡 The app will use a default template instead. Common issues:")
                st.write("- File might be corrupted or password protected")
                st.write("- Template might have unsupported features")
                st.write("- Try using a simpler .pptx template")

                # Reset template data
                generator.template_prs = None
                generator.template_styles = {}

    # Main content area
    col1, col2 = st.columns([2, 1])

    with col1:
        st.header("📝 Your Content")

        # Input text area
        sample_text = """# The Future of Artificial Intelligence

    Artificial Intelligence (AI) is rapidly transforming industries and society. From healthcare to finance, AI-driven solutions are enabling automation, improving decision-making, and unlocking new possibilities.

    ## Key Trends
    - Machine learning and deep learning advancements
    - Natural language processing breakthroughs
    - AI-powered automation in business processes

    ## Challenges
    - Ethical considerations and bias
    - Data privacy and security
    - Workforce displacement and reskilling

    ## Opportunities
    - Personalized healthcare and diagnostics
    - Smart cities and transportation
    - Enhanced customer experiences

    In conclusion, AI holds immense promise, but responsible development and deployment are crucial for maximizing its benefits while minimizing risks."""

        input_text = st.text_area(
            "Paste your content here",
            value=sample_text,
            height=300,
            placeholder="Paste your markdown, prose, or any text content here. The AI will intelligently break it down into slide-worthy content...",
            help="Supports markdown formatting and long-form text"
        )

        # Style guidance
        st.header("🎯 Presentation Style")
        guidance_options = [
            "",
            "investor pitch deck",
            "research summary",
            "sales presentation",
            "technical documentation",
            "educational content",
            "project proposal",
            "quarterly review",
            "product launch"
        ]

        guidance = st.selectbox(
            "Choose presentation style (optional)",
            guidance_options,
            help="This guides the AI on how to structure and tone your presentation"
        )

        if guidance == "":
            custom_guidance = st.text_input(
                "Or enter custom guidance",
                placeholder="e.g., board meeting presentation, customer onboarding, training material"
            )
            if custom_guidance:
                guidance = custom_guidance

    with col2:
        st.header("🚀 Features")

        st.markdown("""
        <div class="feature-box">
            <h4>🤖 AI-Powered Analysis</h4>
            <p>Intelligently structures your content into compelling slides</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="feature-box">
            <h4>🎨 Template Preservation</h4>
            <p>Maintains your brand colors, fonts, and layouts. Custom fonts fallback to system fonts for compatibility.</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="feature-box">
            <h4>🔒 Privacy First</h4>
            <p>No data storage - everything processed securely</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="feature-box">
            <h4>⚡ Multiple AI Providers</h4>
            <p>Choose from OpenAI, Anthropic, or Google</p>
        </div>
        """, unsafe_allow_html=True)

    # Generation button and process
    st.divider()

    if st.button("🚀 Generate Presentation", type="primary", use_container_width=True):
        if not input_text.strip():
            st.error("❌ Please enter some content to convert.")
            return

        if not api_key.strip():
            st.error("❌ Please enter your API key.")
            return

        try:
            with st.spinner("🤖 Analyzing content with AI..."):
                # Create progress bar
                progress_bar = st.progress(0)
                status_text = st.empty()

                # Step 1: Generate structure
                status_text.text("🔍 Analyzing content structure...")
                progress_bar.progress(25)

                prompt = generator.create_prompt(input_text, guidance)
                ai_response = generator.call_ai_api(provider, api_key, prompt)

                # Step 2: Parse response
                status_text.text("📋 Parsing AI response...")
                progress_bar.progress(50)

                structure = generator.parse_ai_response(ai_response)

                # Step 3: Create presentation
                status_text.text("🎨 Creating presentation...")
                progress_bar.progress(75)

                prs = generator.create_presentation(
                    structure, generator.template_prs)

                # Step 4: Finalize
                status_text.text("✅ Finalizing presentation...")
                progress_bar.progress(90)

                # Generate filename
                safe_title = re.sub(
                    r'[^a-z0-9\s]', '', structure.get('title', 'presentation').lower())
                safe_title = re.sub(r'\s+', '_', safe_title)
                filename = f"{safe_title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"

                # Save presentation
                output = io.BytesIO()
                prs.save(output)
                output.seek(0)

                progress_bar.progress(100)
                status_text.text("🎉 Presentation generated successfully!")

                # Success message with details
                st.markdown(f"""
                <div class="success-box">
                    <h4>✅ Success!</h4>
                    <p><strong>Title:</strong> {structure.get('title', 'Generated Presentation')}</p>
                    <p><strong>Slides Created:</strong> {len(structure['slides']) + 1} (including title slide)</p>
                    <p><strong>Template Applied:</strong> {'Yes' if generator.template_prs else 'Default'}</p>
                </div>
                """, unsafe_allow_html=True)

                # Download button
                st.download_button(
                    label="📥 Download Presentation",
                    data=output.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

                # Show presentation structure
                with st.expander("📋 Generated Structure"):
                    st.json(structure)

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.info("💡 Try using a different AI provider or check your API key.")

    # Footer
    st.divider()
    st.markdown("""
    <div style="text-align: center; color: #6b7280; padding: 2rem;">
        <p>Made with ❤️ using Streamlit • Transform your ideas into professional presentations</p>
        <p><small>Your API keys and content are processed securely and never stored</small></p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
