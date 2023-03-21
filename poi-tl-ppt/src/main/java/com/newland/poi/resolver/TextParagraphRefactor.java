package com.newland.poi.resolver;

import com.deepoove.poi.resolver.RunEdge;
import com.newland.poi.xslf.XslfTextParagraphWrapper;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 重构文本段落格式
 *
 * @author xuzhh
 * @since 1.0 2023/3/19
 */
public class TextParagraphRefactor {

    private static final Logger LOG = LoggerFactory.getLogger(TextParagraphRefactor.class);

    private final XslfTextParagraphWrapper paragraph;
    private final List<XSLFTextRun> runs;

    List<Pair<RunEdge, RunEdge>> pairs = new ArrayList<>();

    /**
     * @param paragraph 待处理文本段
     * @param pattern   识别占位符段正则
     */
    public TextParagraphRefactor(XSLFTextParagraph paragraph, Pattern pattern) {
        Objects.requireNonNull(paragraph, "paragraph cannot be null !");
        this.paragraph = new XslfTextParagraphWrapper(paragraph);
        this.runs = paragraph.getTextRuns();
        if (null == runs || runs.isEmpty()) {
            return;
        }

        buildRunEdge(pattern);
    }

    public List<XSLFTextRun> refactorRun() {
        if (pairs.isEmpty()) {
            return Collections.emptyList();
        }

        List<XSLFTextRun> templateRuns = new ArrayList<>();
        int size = pairs.size();
        Pair<RunEdge, RunEdge> runEdgePair;
        for (int n = size - 1; n >= 0; n--) {
            runEdgePair = pairs.get(n);
            RunEdge startEdge = runEdgePair.getLeft();
            RunEdge endEdge = runEdgePair.getRight();
            int startRunPos = startEdge.getRunPos();
            int endRunPos = endEdge.getRunPos();
            int startOffset = startEdge.getRunEdge();
            int endOffset = endEdge.getRunEdge();

            String startText = runs.get(startRunPos).getRawText();
            String endText = runs.get(endRunPos).getRawText();

            // clear the redundant end Run directly
            if (endOffset + 1 >= endText.length() && startRunPos != endRunPos) {
                paragraph.removeTextRun(runs.get(endRunPos));
            } else {
                // split end run, set extra in a run
                String extra = endText.substring(endOffset + 1);
                if (null != extra && !extra.isEmpty()) {
                    if (startRunPos == endRunPos) {
                        // create run and set extra content
                        XSLFTextRun extraRun = paragraph.insertNewTextRunAfter(endRunPos);
                        buildExtra(extraRun, extra);
                    } else {
                        // Set the extra content to the redundant end run
                        buildExtra(runs.get(endRunPos), extra);
                    }
                }
            }

            // clear extra run
            for (int m = endRunPos - 1; m > startRunPos; m--) {
                paragraph.removeTextRun(runs.get(m));
            }

            if (startOffset <= 0) {
                // set the start Run directly
                XSLFTextRun templateRun = runs.get(startRunPos);
                templateRun.setText(startEdge.getTag());
                templateRuns.add(runs.get(startRunPos));
            } else {
                // split start run, set extra in a run
                String extra = startText.substring(0, startOffset);
                XSLFTextRun extraRun = runs.get(startRunPos);
                buildExtra(extraRun, extra);

                XSLFTextRun templateRun = paragraph.insertNewTextRunAfter(startRunPos);
                templateRun.setText(startEdge.getTag());
                templateRuns.add(runs.get(startRunPos + 1));
            }
        }

        return templateRuns;
    }

    private void buildExtra(XSLFTextRun extraRun, String extra) {
        extraRun.setText(extra);
    }

    private void buildRunEdge(Pattern pattern) {
        // find all templates
        Matcher matcher = pattern.matcher(paragraph.getParagraph().getText());
        while (matcher.find()) {
            pairs.add(ImmutablePair.of(new RunEdge(matcher.start(), matcher.group()), new RunEdge(matcher.end(), matcher.group())));
        }
        if (pairs.isEmpty()) {
            return;
        }

        boolean endFlag = false;
        int size = runs.size();
        int cursor = 0;
        int pos = 0;

        // find the run where all templates are located
        Pair<RunEdge, RunEdge> pair = pairs.get(pos);
        RunEdge startEdge = pair.getLeft();
        RunEdge endEdge = pair.getRight();
        int start = startEdge.getAllEdge();
        int end = endEdge.getAllEdge();
        for (int i = 0; i < size; i++) {
            XSLFTextRun run = runs.get(i);
            String text = run.getRawText();
            // empty run
            if (null == text) {
                LOG.warn("found the empty text run, may be produce bug: {}", run);
                cursor += run.toString().length();
                continue;
            }
            LOG.debug(text);
            // The starting position is not enough, the cursor points to the next run
            if (text.length() + cursor < start) {
                cursor += text.length();
                continue;
            }
            // index text
            for (int offset = 0; offset < text.length(); offset++) {
                if (cursor + offset == start) {
                    startEdge.setRunPos(i);
                    startEdge.setRunEdge(offset);
                    startEdge.setText(text);
                }
                if (cursor + offset == end - 1) {
                    endEdge.setRunPos(i);
                    endEdge.setRunEdge(offset);
                    endEdge.setText(text);

                    if (pos == pairs.size() - 1) {
                        endFlag = true;
                        break;
                    }

                    // Continue to calculate the next template
                    pair = pairs.get(++pos);
                    startEdge = pair.getLeft();
                    endEdge = pair.getRight();
                    start = startEdge.getAllEdge();
                    end = endEdge.getAllEdge();
                }
            }
            if (endFlag) {
                break;
            }
            // the cursor points to the next run
            cursor += text.length();
        }

        loggerInfo();
    }

    public void loggerInfo() {
        for (Pair<RunEdge, RunEdge> runEdges : pairs) {
            LOG.debug("[Start]: {}", runEdges.getLeft());
            LOG.debug("[End]: {}", runEdges.getRight());
        }
    }

}
