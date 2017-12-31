package com.moonspirit.pptextract;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.Scanner;
import org.apache.poi.hslf.usermodel.HSLFNotes;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;

public class Main {

    public static void main(String[] args) throws FileNotFoundException, IOException {
        Scanner in = new Scanner(System.in);
        String fileName = in.nextLine();
        FileInputStream fi = new FileInputStream(fileName);
        FileOutputStream fo = new FileOutputStream(new File(fileName + ".txt"));
        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fo, "utf-8"));
        try (HSLFSlideShow ppt = new HSLFSlideShow(fi)) {
            int i = 0;
            for (HSLFSlide slide : ppt.getSlides()) {
                i++;
                HSLFNotes hslfNotes = slide.getNotes();
                if (null == hslfNotes) {
                    continue;
                }
                List<List<HSLFTextParagraph>> paragraph = hslfNotes.getTextParagraphs();
                for (List paragraphList : paragraph) {
                    String tmp = HSLFTextParagraph.getText(paragraphList);
                    if (!tmp.trim().isEmpty()) {
                        bw.write("Page " + i + ":\n");
                        bw.write(tmp.trim());
                        bw.newLine();
                        bw.flush();
                    }
                }
            }
        }
        fi.close();
        fo.close();
    }

}
