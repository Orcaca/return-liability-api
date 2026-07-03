function detailObjectToRow(item) {
  return [
    item["대분류"] || "",
    item["구분"] || "",
    toNumber(item["총액매출"]),
    toNumber(item["당해매출반품"]),
    toNumber(item["1년이상반품"]),
    toNumber(item["1년"]),
    toNumber(item["2년"]),
    toNumber(item["순매출액"]),
    toNumber(item["원가금액"]),
    toNumber(item["원가율"]),

    // K/L = 적용 반품율
    toNumber(item["적용1년반품율"]),
    toNumber(item["적용2년반품율"]),

    // M/N = 당해 반품율
    toNumber(item["당해1년반품율"]),
    toNumber(item["당해2년반품율"]),

    toNumber(item["당해매출기준_반품추정액"]),
    toNumber(item["전기매출기준_반품추정액"]),
    toNumber(item["당해매출기준_원가추정액"]),
    toNumber(item["전기매출기준_원가추정액"]),
    toNumber(item["순충당부채"]),
    ""
  ];
}
