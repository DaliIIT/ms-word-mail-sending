package de.sanvartis.mailword;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class HtmlUtils {


    public static String readFile(String filePath) throws IOException {
        byte[] encoded = new byte[0];
        encoded = Files.readAllBytes(Paths.get(filePath));
        return new String(encoded, "windows-1252");
    }


    static List<String> getImgs(String HTML_CONTENT) throws IOException {
        final String STRING_PATTERN = "['\"]([^'\"]+)";
        final String SRC_PATTERN = "src\\s*=\\s*" + STRING_PATTERN + "['\"]";
        final String IMG_PATTERN = "<img[^>]+" + SRC_PATTERN + "[^>]*>";
        final Pattern stringPattern = Pattern.compile(STRING_PATTERN);
        final Pattern srcPattern = Pattern.compile(SRC_PATTERN);
        final Pattern ImgPattern = Pattern.compile(IMG_PATTERN);
        // ------------------------------------------------------
        Matcher matcher = ImgPattern.matcher(HTML_CONTENT);
        List<String> imgTags = new ArrayList<>();
        while (matcher.find()) {
            final String group = matcher.group();
            imgTags.add(group);
        }

        return imgTags.stream()
                .map(srcPattern::matcher)
                .peek(Matcher::find)
                .map(Matcher::group)
                .map(stringPattern::matcher)
                .peek(Matcher::find)
                .map(stringMatcher -> stringMatcher.group(1))
                .map(s -> s.replaceAll("/", "\\\\"))
                .collect(Collectors.toList());
    }

    static String toUnixPath(String pattern) {
        final String s = pattern.replaceAll("\\\\", "/");
        System.out.println(s);
        return s;
    }

    static String getImageName(String name) {
        return name.replaceAll(".*\\\\", "");
    }
}
