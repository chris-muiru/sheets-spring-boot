package com.muiru.sheets.main.service;

import java.io.IOException;
import java.nio.file.Path;

public interface SheetService {
    Path exportExcelFile() throws IOException;
}
