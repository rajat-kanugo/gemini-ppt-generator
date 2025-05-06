package com.example.gemini_slide_generator;

import com.google.gson.*;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.springframework.http.*;
import org.springframework.web.client.RestTemplate;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Properties;

public class SlideGenerator {

    private static String loadApiKey() throws IOException {
        Properties props = new Properties();
        props.load(new FileInputStream("application.properties"));
        return props.getProperty("GEMINI_API_KEY");
    }

    public static void main(String[] args) {
        try {
            String apiKey = loadApiKey();
            String GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent?key=" + apiKey;

            RestTemplate restTemplate = new RestTemplate();

            // Request input for the slide topic
            String userInput = "Enter your slide topic:. Provide a title slide and 3 content slides.";

            // Construct the request payload
            JsonObject prompt = new JsonObject();
            JsonArray contents = new JsonArray();
            JsonObject part = new JsonObject();
            part.addProperty("text", userInput);
            JsonObject item = new JsonObject();
            item.add("parts", new JsonArray());
            item.getAsJsonArray("parts").add(part);
            contents.add(item);
            prompt.add("contents", contents);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_JSON);
            HttpEntity<String> entity = new HttpEntity<>(prompt.toString(), headers);

            // Make the API request
            ResponseEntity<String> response = restTemplate.postForEntity(GEMINI_API_URL, entity, String.class);
            System.out.println("Gemini API Response: " + response.getBody());

            JsonObject responseJson = JsonParser.parseString(response.getBody()).getAsJsonObject();
            JsonArray candidates = responseJson.getAsJsonArray("candidates");

            if (candidates == null || candidates.size() == 0) {
                throw new RuntimeException("No candidates found in Gemini response");
            }

            // Get the first candidate's content
            JsonObject firstCandidate = candidates.get(0).getAsJsonObject();
            JsonObject content = firstCandidate.getAsJsonObject("content");
            JsonArray parts = content.getAsJsonArray("parts");

            if (parts == null || parts.size() == 0) {
                throw new RuntimeException("No slide content found in Gemini response");
            }

            // Extract slide text
            String slideText = parts.get(0).getAsJsonObject().get("text").getAsString();

            // Generate PowerPoint presentation from the response text
            createPPT(slideText);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createPPT(String slideText) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        Color backgroundColor = new Color(34, 94, 124);

        // Split slides based on paragraph groups
        String[] slideBlocks = slideText.split("(?=## Slide)"); // Keep "## Slide" markers

        for (String block : slideBlocks) {
            block = block.trim();
            if (block.isEmpty()) continue;

            // Further split lines for granular control
            String[] lines = block.split("\n");
            StringBuilder currentSlideContent = new StringBuilder();
            int lineCounter = 0;
            int maxLines = 10; // adjust based on font size and box height

            for (String line : lines) {
                if (lineCounter >= maxLines) {
                    createSlide(ppt, backgroundColor, currentSlideContent.toString());
                    currentSlideContent = new StringBuilder();
                    lineCounter = 0;
                }
                currentSlideContent.append(line).append("\n");
                lineCounter++;
            }

            if (currentSlideContent.length() > 0) {
                createSlide(ppt, backgroundColor, currentSlideContent.toString());
            }
        }

        try (FileOutputStream out = new FileOutputStream("generated_presentation.pptx")) {
            ppt.write(out);
        }

        System.out.println("Presentation created successfully.");
    }



    private static void createSlide(XMLSlideShow ppt, Color backgroundColor, String slideText) {
        XSLFSlide slide = ppt.createSlide();

        // Background color
        slide.getBackground().setFillColor(backgroundColor);

        // Create a larger white textbox
        XSLFTextShape shape = slide.createTextBox();

        // Increase height to accommodate more content (e.g. 500px instead of 400)
        shape.setAnchor(new Rectangle(50, 30, 600, 500)); // x, y, width, height
        shape.setFillColor(Color.WHITE);
        shape.setLineColor(Color.LIGHT_GRAY); // optional border

        // Add text
        XSLFTextParagraph para = shape.addNewTextParagraph();
        para.setLineSpacing(110.0);  // Slightly more spacing between lines

        XSLFTextRun run = para.addNewTextRun();
        run.setText(slideText.trim());
        run.setFontSize(20.0);
        run.setFontColor(Color.BLACK);
        run.setFontFamily("Arial");
    }

    private static void addSlide(XMLSlideShow ppt, String text, Color bgColor) {
        XSLFSlide slide = ppt.createSlide();
        slide.getBackground().setFillColor(bgColor);

        XSLFTextShape shape = slide.createTextBox();
        shape.setAnchor(new Rectangle(50, 50, 600, 400));
        shape.setFillColor(Color.WHITE);
        shape.setLineColor(Color.LIGHT_GRAY);

        XSLFTextParagraph para = shape.addNewTextParagraph();
        XSLFTextRun run = para.addNewTextRun();
        run.setText(text);
        run.setFontSize(20.0);
        run.setFontColor(Color.BLACK);
        run.setFontFamily("Arial");
    }
    }

