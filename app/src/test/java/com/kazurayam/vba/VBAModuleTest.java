package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonProcessingException;
import org.testng.annotations.Test;
import static org.assertj.core.api.Assertions.assertThat;

public class VBAModuleTest {

    @Test
    public void test_getName() {
        VBAModule module = new VBAModule("module1");
        assertThat(module.getName()).isEqualTo("module1");
    }

    @Test
    public void tst_toJson() throws JsonProcessingException {
        VBAModule module = new VBAModule("module1");
        assertThat(module.toJson()).isEqualTo("\"module1\"");
    }
}
