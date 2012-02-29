/*
 * Copyright 2002-2012 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.hillert.transformer;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.integration.Message;
import org.springframework.integration.annotation.Transformer;
import org.springframework.integration.file.FileHeaders;
import org.springframework.integration.support.MessageBuilder;

/**
 * This Spring Integration transformation handler takes the input file, converts
 * the file into a string, converts the file contents into an upper-case string
 * and then sets a few Spring Integration message headers.
 *
 */
public class TransformationHandler {

    /**
     * Actual Spring Integration transformation handler.
     *
     * @param inputMessage Spring Integration input message
     * @return New Spring Integration message with updated headers
     */
    @Transformer
    public Message<byte[]> handleFile(final Message<File> inputMessage) {

        final File inputFile = inputMessage.getPayload();
        final String filename = inputFile.getName();


        final String inputAsString;

        try {
            inputAsString = FileUtils.readFileToString(inputFile);
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }


        ByteArrayOutputStream bout = new ByteArrayOutputStream();

        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Sample Sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow((short)0);
        // Create a cell and put a value in it.
        Cell cell = row.createCell(0);

        cell.setCellValue(inputAsString);

        try {
			wb.write(bout);
		} catch (IOException e) {
			throw new IllegalStateException(e);
		}

        final Message<byte[]> message = MessageBuilder.withPayload(bout.toByteArray())
                      .setHeader(FileHeaders.FILENAME,      filename + ".xls")
                      .setHeader(FileHeaders.ORIGINAL_FILE, inputFile)
                      .setHeader("file_size", inputFile.length())
                      .build();

        return message;
    }

}
