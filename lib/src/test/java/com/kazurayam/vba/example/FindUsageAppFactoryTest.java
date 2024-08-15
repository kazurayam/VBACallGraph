package com.kazurayam.vba.example;

import com.kazurayam.vba.puml.FindUsageApp;
import org.testng.annotations.Test;

import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

public class FindUsageAppFactoryTest {

    @Test
    public void test_createKazurayamSeven() throws IOException {
        FindUsageApp app = FindUsageAppFactory.createKazurayamSeven();
        assertThat(app.size()).isEqualTo(7);
    }
}
