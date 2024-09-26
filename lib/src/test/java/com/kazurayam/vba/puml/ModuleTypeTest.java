package com.kazurayam.vba.puml;

import org.testng.annotations.Test;

import com.kazurayam.vba.puml.VBAModule.ModuleType;

import static org.assertj.core.api.Assertions.assertThat;

public class ModuleTypeTest {

    @Test
    public void testStandard() {
        assertThat(ModuleType.Standard.getFileExtension())
                .isEqualTo(".bas");
    }

    @Test
    public void testClass() {
        assertThat(ModuleType.Class.getFileExtension())
                .isEqualTo(".cls");
    }

    @Test
    public void testDocument() {
        assertThat(ModuleType.Document.getFileExtension())
                .isEqualTo(".doccls");
    }

    @Test
    public void testForm() {
        assertThat(ModuleType.Form.getFileExtension())
                .isEqualTo(".frm");
    }
}
