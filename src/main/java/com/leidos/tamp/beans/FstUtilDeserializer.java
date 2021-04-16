package com.leidos.tamp.beans;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.deser.std.StdDeserializer;
import com.fasterxml.jackson.databind.node.DoubleNode;
import com.fasterxml.jackson.databind.node.IntNode;
import com.leidos.tamp.type.PIECEWISELINEARITEM_TYPE;

import java.io.IOException;

public class FstUtilDeserializer extends StdDeserializer<PIECEWISELINEARITEM_TYPE> {

    public FstUtilDeserializer() {
        this(null);
    }

    public FstUtilDeserializer(Class<?> vc) {
        super(vc);
    }

    @Override
    public PIECEWISELINEARITEM_TYPE deserialize(JsonParser jp, DeserializationContext ctxt)
            throws IOException, JsonProcessingException {
        JsonNode node = jp.getCodec().readTree(jp);
        String name = node.get("Model Name").asText();
        int index = (Integer) ((IntNode) node.get("Index")).numberValue();
        String id = node.get("ID").asText();
        double from = node.get("From FST Hours").doubleValue();
        double to = node.get("To FST Hours").doubleValue();
        double b = node.get("B").doubleValue();
        double m = node.get("M").doubleValue();

        PIECEWISELINEARITEM_TYPE type = new PIECEWISELINEARITEM_TYPE();
        type.setModelName(name);
        type.setFromValue(from);
        type.setToValue(to);
        type.setB(b);
        type.setM(m);

        return type;
    }
}
