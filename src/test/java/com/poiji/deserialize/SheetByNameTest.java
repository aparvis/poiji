package com.poiji.deserialize;

import com.poiji.bind.Poiji;
import com.poiji.deserialize.model.byid.Person;
import com.poiji.deserialize.model.byname.PersonByName;
import com.poiji.option.PoijiOptions;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;

import java.io.File;
import java.util.Arrays;
import java.util.List;

import static com.poiji.util.Data.unmarshallingPersons;
import static org.hamcrest.CoreMatchers.instanceOf;
import static org.junit.Assert.*;

/**
 * Created by ar on 9/03/2018.
 */
@RunWith(Parameterized.class)
public class SheetByNameTest {

    private String path;
    private List<Person> expectedPersonList;
    private Class<?> expectedException;

    public SheetByNameTest(String path, List<Person> expectedPersonList, Class<?> expectedException) {
        this.path = path;
        this.expectedPersonList = expectedPersonList;
        this.expectedException = expectedException;
    }

    @Parameterized.Parameters(name = "{index}: ({0})={1}")
    public static Iterable<Object[]> queries() {
        return Arrays.asList(new Object[][]{
                {"src/test/resources/named_sheets.xlsx", unmarshallingPersons(), null},
                {"src/test/resources/named_sheets.xls", unmarshallingPersons(), null}
        });
    }

    @Test
    public void testRowNumberForXLSXFormatFile() {
        try {
            final PoijiOptions.PoijiOptionsBuilder builder = PoijiOptions.PoijiOptionsBuilder.settings();
            builder.sheetName("Sheet #2");

            List<PersonByName> actualPersons = Poiji.fromExcel(new File(path), PersonByName.class, builder.build());
            assertEquals(expectedPersonList.get(0).getRow(), actualPersons.get(0).getRow());
            assertEquals(expectedPersonList.get(1).getRow(), actualPersons.get(1).getRow());
            assertEquals(expectedPersonList.get(2).getRow(), actualPersons.get(2).getRow());
            assertEquals(expectedPersonList.get(3).getRow(), actualPersons.get(3).getRow());
            assertEquals(expectedPersonList.get(4).getRow(), actualPersons.get(4).getRow());
        } catch (Exception e) {
            if (expectedException == null) {
                fail(e.getMessage());
            } else {
                assertThat(e, instanceOf(expectedException));
            }
        }
    }

}
