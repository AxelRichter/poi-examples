package arichter.examples.apache.poi.sl;

import org.apache.poi.xslf.usermodel.*;
import java.util.List;
import java.util.ArrayList;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
* ReplaceTextInRuns provides a replace method that can be used to replace one text with another, 
* even if the text is broken into multiple pieces of text within the body paragraph. 
*
* @author Axel Richter
* 
* @see arichter.examples.apache.poi.wp.ReplaceTextInRuns
* */
public class ReplaceTextInRuns {
    
    private static final Logger LOG = LogManager.getLogger(ReplaceTextInRuns.class);
    
    /**
    * Constructor not used
    */
    public ReplaceTextInRuns() {
    }

    
    private String s1EndsWithLeftPartOfS2(String s1, String s2) {
        String overlap = "";
        for (int i = 0; i < s2.length(); i++) {
            if (s1.endsWith(s2.substring(0, i+1))) {
                overlap = s2.substring(0, i+1);
            }
        }
        return overlap;
    }

    private String s1StartsWithRightPartOfS2(String s1, String s2) {
        String overlap = "";
        for (int i = s2.length() - 1; i >= 0; i--) {
            if (s1.startsWith(s2.substring(i))) {
                overlap = s2.substring(i);
            }
        }
        return overlap;
    }

    private String replaceAtEnd(String string, String toReplace, String replacement) {
        String result = string;
        int posAtEnd = string.length() - toReplace.length();
        int pos = string.lastIndexOf(toReplace);
        if (pos == posAtEnd) {
            result = string.substring(0, pos) + replacement;
        }
        return result;
    }

    private String replaceAtBegin(String string, String toReplace, String replacement) {
        String result = string;
        if (string.startsWith(toReplace)) {
            result = replacement + string.substring(toReplace.length());
        }
        return result;
    }

    /**
    * A method that can be used to replace one text with another,
    * even if the text is broken into multiple pieces of text within the body paragraph.
    * 
    * @param paragraph the text paragraph which probably contains String marker
    * @param marker the text string that shall be replaced
    * @param newText the text string that shall be the replacement
    * @return a {@link java.util.List} containing all text runs which contains the replaced text
    */
    public List<XSLFTextRun> replace(XSLFTextParagraph paragraph, String marker, String newText) {
        XSLFTextRun run;
        String text;
        XSLFTextRun prevRun;
        String prevText;
        XSLFTextRun nextRun;
        List<XSLFTextRun> nextRuns;
        XSLFTextRun midRun;
        String nextText;
        String leftOverlap;
        String rightOverlap;
        String midPartText;
        String partOfMarker;

        List<XSLFTextRun> resultRuns = new ArrayList<XSLFTextRun>();
        List<XSLFTextRun> runs;
  
        if (paragraph.getText().contains(marker)) {
            runs = paragraph.getTextRuns();
            for (int i = 0; i < runs.size(); i++) {
                run = runs.get(i);
                text = run.getRawText(); if (text == null) text = "";
                if (text.contains(marker)) { //
                    run.setText(text.replace(marker, newText));
                    resultRuns.add(run);
                } else if (i > 0) {
                    prevRun = runs.get(i - 1);
                    prevText = prevRun.getRawText(); if (prevText == null) prevText = "";
                    leftOverlap = s1EndsWithLeftPartOfS2(prevText, marker);
                    if (!"".equals(leftOverlap)) {
                        if ((leftOverlap + text).startsWith(marker)) {
                            partOfMarker = marker.substring(leftOverlap.length());
                            prevRun.setText(replaceAtEnd(prevText, leftOverlap, newText));
                            resultRuns.add(prevRun);
                            run.setText(replaceAtBegin(text, partOfMarker, ""));
                        } else {
                            nextRuns = new ArrayList<XSLFTextRun>();
                            rightOverlap = "";
                            for(int j = i ; j < runs.size(); j++) {
                                nextRun = runs.get(j);
                                nextRuns.add(nextRun);
                                nextText = nextRun.getRawText(); if (nextText == null) nextText = "";
                                rightOverlap = s1StartsWithRightPartOfS2(nextText, marker);
                                if (!"".equals(rightOverlap)) break;
                            }
                            if (!"".equals(rightOverlap) && nextRuns.size() > 0) {
                                midPartText = "";
                                for (int k = 0; k < nextRuns.size() - 1; k++) {
                                    midRun = nextRuns.get(k);
                                    String midRunText = midRun.getRawText(); if (midRunText == null) midRunText = "";
                                    midPartText += midRunText;
                                }
                                if ((leftOverlap + midPartText + rightOverlap).equals(marker)) {
                                    prevRun.setText(replaceAtEnd(prevText, leftOverlap, newText));
                                    resultRuns.add(prevRun);
                                    for (int k = 0; k < nextRuns.size() - 1; k++) {
                                        midRun = nextRuns.get(k);
                                        midRun.setText("");
                                    }
                                    nextRun = nextRuns.get(nextRuns.size() - 1);
                                    nextText = nextRun.getRawText(); if (nextText == null) nextText = "";
                                    nextRun.setText(replaceAtBegin(nextText, rightOverlap, ""));
                                }
                            }
                        }
                    }
                }
            }
        }
        return resultRuns;
    }
}
