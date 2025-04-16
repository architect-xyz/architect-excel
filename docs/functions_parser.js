fetch('functions.json')
.then(response => response.json())
.then(data => {
  const container = document.getElementById('functions-container');

  data.functions.forEach(fn => {
    const section = document.createElement('section');
    section.id = fn.name; // for deep-linking

    const title = document.createElement('h2');
    title.textContent = fn.name;
    section.appendChild(title);

    const description = document.createElement('p');
    description.textContent = fn.description || 'No description available.';
    section.appendChild(description);

    // Parameters
    if (fn.parameters && fn.parameters.length > 0) {
      const paramHeader = document.createElement('h3');
      paramHeader.textContent = 'Parameters';
      section.appendChild(paramHeader);

      const paramList = document.createElement('ul');
      fn.parameters.forEach(param => {
        const item = document.createElement('li');
        item.innerHTML = `<strong>${param.name}</strong> (${param.type}): ${param.description}`;
        paramList.appendChild(item);
      });
      section.appendChild(paramList);
    } else {
        const noParams = document.createElement('p');
        noParams.textContent = 'No parameters.';
        section.appendChild(noParams);
    }

    // Result
    if (fn.result) {
      const result_type = document.createElement('p');
      const dim = fn.result.dimensionality ? `${fn.result.dimensionality} ` : '';
      result_type.innerHTML = `<strong>Return Type:</strong> ${dim}${fn.result.type}`;
      section.appendChild(result_type);
    }

    container.appendChild(section);
  });
})
.catch(error => {
  console.error('Failed to load functions.json:', error);
});