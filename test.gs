function myFunction(query_object) {
  var param_value_list = Object.entries(query_object);
  var param_strings = param_value_list.map(([param, value]) => `${param}=${value}`);
  var query_string = param_strings.join('&');
  return '?' + query_string;
}

function main()
{
  var x = {"jikael": 6, "ella": 4};
  console.log(myFunction(x))
}
