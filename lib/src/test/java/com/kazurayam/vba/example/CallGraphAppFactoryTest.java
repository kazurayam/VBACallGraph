package com.kazurayam.vba.example;

import com.kazurayam.vba.puml.CallGraphApp;
import org.testng.annotations.Test;

import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

public class CallGraphAppFactoryTest {

    @Test
    public void test_createKazurayamSeven() throws IOException {
        CallGraphApp app = CallGraphAppFactory.createKazurayamSeven();
        assertThat(app.size()).isEqualTo(6);
    }

    @Test
    public void test_createKazurayamSevenPlus() throws IOException {
        CallGraphApp app = CallGraphAppFactory.createKazurayamSevenPlus();
        assertThat(app.size()).isEqualTo(7);
    }
}
