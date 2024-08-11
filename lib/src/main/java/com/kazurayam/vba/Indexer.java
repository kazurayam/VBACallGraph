package com.kazurayam.vba;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import java.util.List;
import java.util.SortedSet;
import java.util.TreeSet;

/**
 * The Indexer class analyses a set of SensibleWorkbook objects
 * to find out a set of ProcedureReferences found amongst the
 * workbooks. The Indexer supports 2 ways of indexing.
 * - Partial indexing; accept a single FullyQualifiedProcedureId as a referee,
 * scan all the rest if the other procedure refers to the referee. Will cache
 * the scanning result, and return a set of ProcedureReferences found.
 * - Whole indexing; do the partial index for all of candidate referees as one batch.
 */
public class Indexer {

    private final List<SensibleWorkbook> workbooks;
    private final Set<ProcedureReference> memo;

    private static final ObjectMapper mapper;
    static {
        mapper = new ObjectMapper();
        SimpleModule module = new SimpleModule();
        module.addSerializer(Indexer.class,
                new IndexerSerializer());
        module.addSerializer(ProcedureReference.class,
                new ProcedureReference.ProcedureReferenceSerializer());
        module.addSerializer(FullyQualifiedProcedureId.class,
                new FullyQualifiedProcedureId.FullyQualifiedProcedureIdSerializer());
        mapper.registerModule(module);
    }

    public Indexer() {
        workbooks = new ArrayList<>();
        memo = new HashSet<>();
    }

    public void add(SensibleWorkbook workbook) {
        this.workbooks.add(workbook);
    }

    public List<SensibleWorkbook> getWorkbooks() {
        return workbooks;
    }

    /**
     * This method is the core of this project.
     */
    public SortedSet<ProcedureReference> findReferenceTo(FullyQualifiedProcedureId referee) {
        SortedSet<ProcedureReference> scanResultByReferee = this.scanMemoByReferee(referee);
        if (scanResultByReferee.isEmpty()) {
            SortedSet<ProcedureReference> crossReferences = xref(workbooks, referee);
            memo.addAll(crossReferences);
            return crossReferences;
        } else {
            return scanResultByReferee;
        }
    }

    SortedSet<ProcedureReference> scanMemoByReferee(FullyQualifiedProcedureId referee) {
        SortedSet<ProcedureReference> found = new TreeSet<>();
        for (ProcedureReference ref : memo) {
            if (ref.getReferee().equals(referee)) { found.add(ref); }
        }
        return found;
    }

    /**
     * The magic spell
     */
    SortedSet<ProcedureReference> xref(List<SensibleWorkbook> workbooks, FullyQualifiedProcedureId referee) {
        SortedSet<ProcedureReference> result = new TreeSet<>();
        for (SensibleWorkbook workbook : workbooks) {
            for (String moduleName : workbook.getModules().keySet()) {
                VBAModule module = workbook.getModule(moduleName);
                VBASource source = module.getVBASource();
                for (VBAProcedure procedure : module.getProcedures()) {
                    if (referee.getWorkbook().equals(workbook) &&
                            referee.getModule().equals(module) &&
                            referee.getProcedure().equals(procedure)) {
                        break; // we won't scan the VBA source of the referee itself
                    }
                    // let's scan the VBA source if it mentions the referee
                    List<VBASourceLine> linesFound = source.find(procedure.getName());
                    if (!linesFound.isEmpty()) {
                        // referrer(s) to this referee found!
                        for (VBASourceLine line : linesFound) {
                            FullyQualifiedProcedureId referrer =
                                    new FullyQualifiedProcedureId(workbook, module, procedure);
                            ProcedureReference reference =
                                    new ProcedureReference(referrer, referee);
                            result.add(reference);
                        }
                    }
                }
            }
        }
        return result;
    }

    public Iterator<ProcedureReference> iterator() {
        return memo.iterator();
    }

    @Override
    public String toString() {
        // pretty printed
        try {
            Object json = mapper.readValue(this.toJson(), Object.class);
            return mapper.writerWithDefaultPrettyPrinter().writeValueAsString(json);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }

    public String toJson() throws JsonProcessingException {
        return mapper.writeValueAsString(this);
    }

    /**
     *
     */
    private static class IndexerSerializer extends StdSerializer<Indexer> {
        public IndexerSerializer() { this(null); }
        public IndexerSerializer(Class<Indexer> t) {
            super(t);
        }
        @Override
        public void serialize(
                Indexer indexer, JsonGenerator jgen, SerializerProvider provider)
                throws IOException {
            jgen.writeStartObject();                             //{
            jgen.writeArrayFieldStart("workbooks");      // "workbooks":[
            for (SensibleWorkbook wb : indexer.getWorkbooks()) {
                jgen.writeString(wb.getId());
            }
            jgen.writeEndArray();                                //  ],
            jgen.writeArrayFieldStart("procedureReferences");     //    "workbooks": [
            Iterator<ProcedureReference> iter = indexer.iterator();
            while(iter.hasNext()) {
                ProcedureReference reference = iter.next();
                jgen.writeObject(reference);                            //      { ... },
            }
            jgen.writeEndArray();                                //    ]
            jgen.writeEndObject();                               //}
        }
    }
}
