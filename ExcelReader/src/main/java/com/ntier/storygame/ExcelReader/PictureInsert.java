package com.ntier.storygame.ExcelReader;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javafx.embed.swing.SwingFXUtils;
import javafx.scene.image.WritableImage;

public class PictureInsert {

//	public static void main(String[] args) throws FileNotFoundException {
//		Image img = new Image(new FileInputStream("C:\\Users\\Grayson\\Desktop\\storygame\\testing\\card.png"));
//		WritableImage wImg = new WritableImage(img.getPixelReader(), 600, 400);
//		System.out.println(insertPicture("C:\\Users\\Grayson\\Desktop\\storygame\\testing\\testing.docx", wImg));
//
//	}

	public static boolean insertPicture(String wordDocPath, WritableImage snap) {
		BufferedImage tempImg = SwingFXUtils.fromFXImage(snap, null);
		byte[] imageInByte;
		try {
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			ImageIO.write(tempImg, "png", baos);
			baos.flush();
			imageInByte = baos.toByteArray();
			baos.close();
			File docFile = new File(wordDocPath);
			XWPFDocument document;
			if (docFile.length() > 0)
				document = new XWPFDocument(new FileInputStream(docFile));
			else
				document = new XWPFDocument();
			XWPFRun run = document.createParagraph().createRun();
			run.addCarriageReturn();
			run.addPicture(new ByteArrayInputStream(imageInByte), Document.PICTURE_TYPE_PNG, "", 
					Units.pixelToEMU(600),Units.pixelToEMU(400));
			document.write(new FileOutputStream(wordDocPath));
			document.close();
		} catch (IOException | InvalidFormatException e) {
			e.printStackTrace();
			return false;
		}
		return true;

	}

}
